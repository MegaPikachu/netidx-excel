use anyhow::Result;
use netidx::path::Path;
use netidx::subscriber::Value;

use netidx::subscriber::{Dval, Subscriber};
use std::collections::HashMap;
use tokio::runtime::Runtime;
use std::sync::Mutex;

use netidx::config::Config;
use netidx::resolver_client::DesiredAuth;

#[repr(i16)]
pub enum Send_result {
    Maybe_sent = -2,
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
    pub fn new() -> ExcelNetidxWriter {
        let cfg = Config::load_default().unwrap();
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .thread_name("netidx-writer")
            .build()
            .unwrap();
        let subscriber =
            rt.block_on(
                async move { Subscriber::new(cfg, DesiredAuth::Anonymous).unwrap() },
            );
        let subscribe_writer = SubscribeWriter::new(subscriber);
        ExcelNetidxWriter {
            subscribe_writer,
            rt,
        }
    }

    pub fn send(&self, path: &str, value: Value) -> Send_result {
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

    fn write(&self, path: Path, value: Value) -> Send_result {
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
            Send_result::Sent
        } else {
            Send_result::Maybe_sent
        }
    }
}
