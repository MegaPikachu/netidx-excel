use netidx::subscriber::Value;
use netidx::{
    path::Path,
};
use anyhow::Result;

use netidx::{
    subscriber::{Dval, Subscriber},
};
use std::{
    collections::HashMap,
};

use netidx::{
    config::Config,
};
use tokio_stream::StreamExt;
use netidx::{
    resolver_client::DesiredAuth,
};

pub struct ExcelNetidxWriter {
    events_tx: tokio::sync::mpsc::Sender<WriterEvents>,
}

impl ExcelNetidxWriter {
    pub fn new() -> ExcelNetidxWriter{
        let (events_tx, events_rx) = tokio::sync::mpsc::channel::<WriterEvents>(1_024);
        let cfg = Config::load_default().unwrap();
        let rt = tokio::runtime::Builder::new_multi_thread()
                            .enable_all()
                            .thread_name("excel-publisher")
                            .build()
                            .unwrap();
        let subscriber = rt.block_on(async move {
            Subscriber::new(cfg, DesiredAuth::Anonymous).unwrap()
        });
        std::thread::spawn(move ||{
            rt.block_on(async move {
                let mut subscribe_writer = SubscribeWriter::new(subscriber);
                let mut events_rx = Box::pin(tokio_stream::wrappers::ReceiverStream::new(events_rx));
                loop {
                    match events_rx.next().await {
                        Some(WriterEvents::Write(path, value)) => {
                            let path = Path::from(path);
                            subscribe_writer.write(&path, value);
                        },
                        None => {},
                    }
                }
                
            });
        });
        ExcelNetidxWriter {events_tx}
    }

    pub fn send(&self, path: String, value: Value) -> Result<()>{
        let rt = tokio::runtime::Builder::new_multi_thread()
                            .enable_all()
                            .thread_name("excel-publisher")
                            .build()
                            .unwrap();
        match rt.block_on(async move {
            self.events_tx.send(WriterEvents::Write(path, value)).await
        }){
            Ok(_) => Ok(()),
            Err(err) => Err(err.into()),
        }
        
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
        SubscribeWriter {
            subscriber,
            subscriptions: HashMap::new(),
        }
    }

    fn write(&mut self, path: &Path, value: Value){
        let dv = self.add_subscription(&path);
        if !dv.write(value){
            eprintln!(
                "WARNING: {} queued writes to {}",
                dv.queued_writes(),
                path
            )
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