use netidx::path::Path;
use netidx::subscriber::Value;

use log::info;
use netidx::config::Config;
use netidx::resolver_client::DesiredAuth;
use netidx::subscriber::{Dval, Subscriber};
use std::collections::HashMap;
use std::sync::Mutex;
use tokio::runtime::Runtime;

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

pub struct ExcelNetidxWriter {
    subscribe_writer: SubscribeWriter,
}

impl ExcelNetidxWriter {
    pub fn new() -> anyhow::Result<ExcelNetidxWriter> {
        let cfg = Config::load_default()?;
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .thread_name("netidx-writer")
            .build()?;
        let subscriber =
            rt.block_on(async move { Subscriber::new(cfg, DesiredAuth::Anonymous) })?;
        let subscribe_writer = SubscribeWriter::new(subscriber, rt);
        Ok(ExcelNetidxWriter { subscribe_writer })
    }

    pub fn send(&self, path: &str, value: Value) -> SendResult {
        let path = Path::from_str(path);
        self.subscribe_writer.write_with_recipt(path, value)
    }

    pub fn refresh_path(&self, path: &str) -> SendResult {
        let path = Path::from_str(path);
        self.subscribe_writer.refresh_path(path);
        SendResult::Sent
    }

    pub fn refresh_all(&self) -> SendResult {
        self.subscribe_writer.refresh_all();
        SendResult::Sent
    }
}

lazy_static::lazy_static! {
    static ref SUBSCRIPTIONS: Mutex<HashMap<Path, Dval>>= Mutex::new(HashMap::new());
    static ref WRITER_WAIT_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(3);
}

struct SubscribeWriter {
    subscriber: Subscriber,
    rt: Runtime,
}

impl SubscribeWriter {
    fn new(subscriber: Subscriber, rt: Runtime) -> Self {
        SubscribeWriter { subscriber, rt }
    }

    fn write(&self, path: Path, value: Value) -> SendResult {
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        if match subscriptions.get(&path) {
            Some(sub) => sub.write(value),
            None => {
                info!("Writer: start subscribe path {}", path.clone());
                let sub = self.subscriber.subscribe(path.clone());
                let result = sub.write(value);
                subscriptions.insert(path, sub);
                result
            }
        } {
            SendResult::Sent
        } else {
            SendResult::MaybeSent
        }
    }

    fn write_with_recipt(&self, path: Path, value: Value) -> SendResult {
        async fn wait_result(
            result: futures::channel::oneshot::Receiver<Value>,
            path: &Path,
        ) {
            match tokio::time::timeout(*WRITER_WAIT_TIMEOUT, result).await {
                Ok(result) => {
                    match result {
                        Ok(_) => {}
                        Err(err) => {
                            log::warn!(
                        "Writer: write error for {}, try to refresh the connection, err = {}",
                        path,
                        &err
                    );
                            let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
                            subscriptions.remove(path);
                        }
                    };
                }
                Err(err) => {
                    log::warn!(
                        "Writer: write timeout for {}, try to refresh the connection, err = {}",
                        path,
                        &err
                    );
                    let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
                    subscriptions.remove(path);
                }
            };
        }

        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        let path_copy = path.clone();
        match subscriptions.get(&path) {
            Some(sub) => {
                let result = sub.write_with_recipt(value);
                self.rt.spawn(async move {
                    wait_result(result, &path_copy).await;
                });
            }
            None => {
                let sub = self.subscriber.subscribe(path.clone());
                let result = sub.write_with_recipt(value);
                self.rt.spawn(async move {
                    wait_result(result, &path_copy).await;
                });
                subscriptions.insert(path, sub);
            }
        };
        SendResult::MaybeSent
    }

    // A temp solution for manual refresh
    fn refresh_path(&self, path: Path) {
        info!("Writer: Refresh Path {}", &path);
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        subscriptions.remove(&path);
    }

    // A temp solution for manual refresh
    fn refresh_all(&self) {
        info!("Writer: Refresh all paths");
        let mut subscriptions = SUBSCRIPTIONS.lock().unwrap();
        subscriptions.clear();
    }
}
