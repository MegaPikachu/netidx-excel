use netidx::path::Path;
use netidx::subscriber::Value;
use tokio::sync::mpsc::UnboundedSender;

use log::info;
use netidx::config::Config;
use netidx::resolver_client::DesiredAuth;
use netidx::subscriber::{Dval, Subscriber};
use std::collections::HashMap;
use std::sync::Mutex;
use std::time::{Duration, SystemTime};
use tokio::runtime::Runtime;
use tokio::sync::mpsc::unbounded_channel;

#[repr(i16)]
pub enum SendResult {
    MaybeSent = -2,
    Sent = -1,
    ExcelErrorNull = 0,
    ExcelErrorDiv0 = 7,
    ExcelErrorValue = 15,
    ExcelErrorRef = 23,
    ExcelErrorName = 29,
    ExcelErrorNum = 36,
    ExcelErrorNA = 42,
    ExcelErrorGettingData = 43,
}

#[repr(i16)]
pub enum RequestType {
    Async = 0,
    Retry = 1,
}

pub struct ExcelNetidxWriter {
    subscriber: Subscriber,
    rt: Runtime,
    tx: UnboundedSender<(Path, Value, SystemTime)>,
}

impl ExcelNetidxWriter {
    pub fn new() -> anyhow::Result<ExcelNetidxWriter> {
        log::info!("Init ExcelNetidxWriter, version = 3.11");
        let cfg = Config::load_default()?;
        let (tx, mut rx) = unbounded_channel::<(Path, Value, SystemTime)>();
        let tx_copy = tx.clone();
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .thread_name("netidx-writer")
            .build()?;
        let subscriber =
            rt.block_on(async move { Subscriber::new(cfg, DesiredAuth::Anonymous) })?;
        let subscriber_copy = subscriber.clone();

        std::thread::spawn(move || {
            let rt = tokio::runtime::Builder::new_multi_thread()
                .enable_all()
                .thread_name("netidx-writer")
                .build()
                .unwrap();

            rt.block_on(async move {
                let mut path_send_time: HashMap<Path, std::time::SystemTime> =
                    HashMap::new();
                while let Some((path, value, send_time)) = rx.recv().await {
                    if path_send_time.contains_key(&path)
                        && send_time <= *path_send_time.get(&path).unwrap()
                    {
                        log::debug!("Skip sending old value, path = {}", &path);
                        continue;
                    } else {
                        path_send_time.insert(path.clone(), send_time);
                    }
                    Self::write_with_recipt_retry(
                        &subscriber_copy,
                        path,
                        value,
                        &tx_copy,
                        send_time,
                    );
                }
            });
        });

        Ok(ExcelNetidxWriter { subscriber, rt, tx })
    }

    // CR-soon ttang: Now I keep this interface the same as before, maybe its better to move it into a separate thread, just as send_retry.
    /// This function doesn't promise data is sent to netidx, but the performance is better.
    pub fn send_async(&self, path: &str, value: Value) -> SendResult {
        let path = Path::from_str(path);
        self.write_with_recipt_async(path, value)
    }

    /// If write_with_recipt functions knows a connection error, it will retry sending only latest data of the same path to netidx.
    pub fn send_retry(&self, path: &str, value: Value) -> SendResult {
        let path = Path::from_str(path);
        self.tx.send((path, value, std::time::SystemTime::now())).unwrap();
        SendResult::MaybeSent
    }

    pub fn refresh_path(&self, path: &str) -> SendResult {
        let path = Path::from_str(path);
        info!("Writer: Refresh Path {}", &path);
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        subscriptions.remove(&path);
        SendResult::Sent
    }

    pub fn refresh_all(&self) -> SendResult {
        info!("Writer: Refresh all paths");
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        subscriptions.clear();
        SendResult::Sent
    }

    async fn wait_result(
        result: futures::channel::oneshot::Receiver<Value>,
        path: &Path,
    ) -> bool {
        match tokio::time::timeout(*WRITER_WAIT_TIMEOUT, result).await {
            Ok(result) => match result {
                Ok(_) => true,
                Err(err) => {
                    log::warn!("Writer: write error for {}, err = {}", path, &err);
                    let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
                    subscriptions.remove(path);
                    false
                }
            },
            Err(err) => {
                log::warn!(
                        "Writer: write timeout for {}, try to refresh the connection, err = {}",
                        path,
                        &err
                    );
                let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
                subscriptions.remove(path);
                false
            }
        }
    }

    fn write_with_recipt_async(&self, path: Path, value: Value) -> SendResult {
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        let path_copy = path.clone();
        match subscriptions.get(&path) {
            Some(sub) => {
                let result = sub.write_with_recipt(value);
                self.rt.spawn(async move {
                    Self::wait_result(result, &path_copy).await;
                });
            }
            None => {
                let sub = self.subscriber.subscribe(path.clone());
                let result = sub.write_with_recipt(value);
                self.rt.spawn(async move {
                    Self::wait_result(result, &path_copy).await;
                });
                subscriptions.insert(path, sub);
            }
        };
        SendResult::MaybeSent
    }

    fn write_with_recipt_retry(
        subscriber: &Subscriber,
        path: Path,
        value: Value,
        tx: &UnboundedSender<(Path, Value, SystemTime)>,
        send_time: SystemTime,
    ) -> SendResult {
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        let path_copy = path.clone();
        match subscriptions.get(&path) {
            Some(sub) => {
                let result = sub.write_with_recipt(value.clone());
                let tx2 = tx.clone();
                tokio::spawn(async move {
                    if !Self::wait_result(result, &path_copy).await {
                        tokio::time::sleep(*RETRY_INTERVAL).await;
                        log::debug!("Resend path {}, with dval", &path_copy);
                        tx2.send((path, value, send_time)).unwrap();
                    }
                });
            }
            None => {
                let sub = subscriber.subscribe(path.clone());
                let result = sub.write_with_recipt(value.clone());
                subscriptions.insert(path.clone(), sub);
                let tx2 = tx.clone();
                tokio::spawn(async move {
                    if !Self::wait_result(result, &path_copy).await {
                        tokio::time::sleep(*RETRY_INTERVAL).await;
                        log::debug!("Resend path {}, without dval", &path_copy);
                        tx2.send((path, value, send_time)).unwrap();
                    }
                });
            }
        };
        SendResult::Sent
    }
}

lazy_static::lazy_static! {
    static ref SUBSCRIPTIONS: Mutex<HashMap<Path, Dval>>= Mutex::new(HashMap::new());
    static ref WRITER_WAIT_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(3);
    static ref RETRY_INTERVAL: std::time::Duration = std::time::Duration::from_secs(1);
}
