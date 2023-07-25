use anyhow::Result;
use netidx::path::Path;
use netidx::subscriber::Value;

use netidx::subscriber::{Dval, Subscriber};
use std::collections::HashMap;
use tokio::runtime::Runtime;

use netidx::config::Config;
use netidx::resolver_client::DesiredAuth;
use tokio_stream::StreamExt;

pub struct ExcelNetidxWriter {
    events_tx: tokio::sync::mpsc::Sender<WriterEvents>,
    rt: Runtime,
}

impl ExcelNetidxWriter {
    pub fn new() -> ExcelNetidxWriter {
        let (events_tx, events_rx) = tokio::sync::mpsc::channel::<WriterEvents>(1_024);
        let cfg = Config::load_default().unwrap();
        let rt = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .thread_name("excel-publisher")
            .build()
            .unwrap();
        let subscriber =
            rt.block_on(
                async move { Subscriber::new(cfg, DesiredAuth::Anonymous).unwrap() },
            );
        std::thread::spawn(move || {
            rt.block_on(async move {
                let mut subscribe_writer = SubscribeWriter::new(subscriber);
                let mut events_rx =
                    Box::pin(tokio_stream::wrappers::ReceiverStream::new(events_rx));
                loop {
                    match events_rx.next().await {
                        Some(WriterEvents::Write(path, value)) => {
                            let path = Path::from(path);
                            subscribe_writer.write(&path, value);
                        }
                        None => {}
                    }
                }
            });
        });
        ExcelNetidxWriter {
            events_tx,
            rt: tokio::runtime::Builder::new_multi_thread()
                .enable_all()
                .thread_name("excel-publisher")
                .build()
                .unwrap(),
        }
    }

    pub fn send(&self, path: String, value: Value) -> Result<()> {
        let events_tx = self.events_tx.clone();
        self.rt
            .spawn(async move { events_tx.send(WriterEvents::Write(path, value)).await });
        Ok(())
    }
}

pub enum WriterEvents {
    Write(String, Value),
}

struct SubscribeWriter {
    subscriber: Subscriber,
    subscriptions: HashMap<Path, Dval>,
}

impl SubscribeWriter {
    fn new(subscriber: Subscriber) -> Self {
        SubscribeWriter { subscriber, subscriptions: HashMap::new() }
    }

    fn write(&mut self, path: &Path, value: Value) {
        let dv = self.add_subscription(&path);
        if !dv.write(value) {
            eprintln!("WARNING: {} queued writes to {}", dv.queued_writes(), path)
        };
    }

    fn add_subscription(&mut self, path: &Path) -> &Dval {
        let subscriptions = &mut self.subscriptions;
        let subscriber = &self.subscriber;
        subscriptions.entry(path.clone()).or_insert_with(|| {
            let s = subscriber.subscribe(path.clone());
            s
        })
    }
}
