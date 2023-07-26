use netidx::path::Path;
use netidx::subscriber::Value;

use netidx::subscriber::{Dval, Subscriber};
use std::collections::HashMap;
use tokio::runtime::Runtime;
use std::sync::Mutex;

use netidx::config::Config;
use netidx::resolver_client::DesiredAuth;

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
    ExcelErrorGettingData = 43
}

pub struct ExcelNetidxWriter {
    subscribe_writer: SubscribeWriter,
    rt: Runtime,
}

impl ExcelNetidxWriter {
    pub fn new() -> anyhow::Result<ExcelNetidxWriter> {
        let cfg = Config::load_default()?;
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .thread_name("netidx-writer")
            .build()?;
        let subscriber =
            rt.block_on(
                async move { Subscriber::new(cfg, DesiredAuth::Anonymous) },
            )?;
        let subscribe_writer = SubscribeWriter::new(subscriber);
        Ok(ExcelNetidxWriter { subscribe_writer, rt })
    }

    pub fn send(&self, path: &str, value: Value) -> SendResult {
        let path = Path::from_str(path);
        self.subscribe_writer.write(path, value)
    }
}

struct SubscribeWriter {
    subscriber: Subscriber,
    subscriptions: Mutex<HashMap<Path, Dval>>,
}

impl SubscribeWriter {
    fn new(subscriber: Subscriber) -> Self {
        SubscribeWriter { subscriber, subscriptions: HashMap::new().into() }
    }

    fn write(&self, path: Path, value: Value) -> SendResult {
        let mut subscriptions = self.subscriptions.lock().unwrap();
        if match subscriptions.get(&path) {
            Some(sub) => sub.write(value),
            None => {
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
}
