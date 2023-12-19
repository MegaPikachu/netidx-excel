pub mod dispatch;
pub mod glue;
pub mod interface;
pub mod variant;

use anyhow::Result;
use dirs;
use log::LevelFilter;
use once_cell::sync::Lazy;
use simplelog;
use std::{
    default::Default,
    fs::{self, File},
    path::PathBuf,
};

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub enum Auth {
    Anonymous,
    Kerberos,
    Tls,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct Config {
    pub log_level: LevelFilter,
    #[serde(default)]
    pub auth_mechanism: Option<Auth>,
}

impl Default for Config {
    fn default() -> Self {
        Config { log_level: LevelFilter::Off, auth_mechanism: None }
    }
}

fn clean_old_logs(folder_path: PathBuf) -> Result<()> {
    let today = chrono::Utc::now();
    let one_week_ago = today - chrono::Duration::weeks(1);
    for file in fs::read_dir(folder_path)? {
        let file_path = file?.path();
        if let Some(file_name) = file_path.file_name() {
            let file_name_str = file_name.to_string_lossy();
            if file_name_str.starts_with("log.")
                && file_name_str.ends_with(".txt")
                && file_name_str.len() == 28
            {
                let date_str = &file_name_str[4..23];
                let date = match chrono::DateTime::parse_from_str(
                    &format!("{date_str}+0000"),
                    "%Y-%m-%dT%H-%M-%S%z",
                ) {
                    Ok(date) => date,
                    Err(_) => {
                        continue;
                    } // ignore invalid filenames
                };
                if date < one_week_ago {
                    log::info!("Deleting file: {:?}", file_path);
                    fs::remove_file(file_path)?;
                }
            }
        }
    }
    Ok(())
}

fn load_config_and_init_log() -> Result<Config> {
    let path = match dirs::config_dir() {
        Some(d) => d,
        None => match dirs::home_dir() {
            Some(d) => d,
            None => PathBuf::from("\\"),
        },
    };
    let base = path.join("netidx-excel");
    let log_base = base.join("logs");
    fs::create_dir_all(log_base.clone())?;
    let config_file = base.join("config.json");
    let now = chrono::Utc::now().format("%Y-%m-%dT%H-%M-%S").to_string();

    let log_file = log_base.join(format!("log.{}Z.txt", now));
    if !config_file.exists() {
        fs::write(&*config_file, &serde_json::to_string_pretty(&Config::default())?)?;
    }
    let config: Config = serde_json::from_str(&fs::read_to_string(config_file.clone())?)?;
    let log = File::create(log_file)?;
    simplelog::WriteLogger::init(config.log_level, simplelog::Config::default(), log)?;
    if let Err(err) = clean_old_logs(log_base.clone()) {
        log::error!("Clean old logs error, err msg = {}", err);
    };
    Ok(config)
}

pub static CONFIG: Lazy<Config> = Lazy::new(|| match load_config_and_init_log() {
    Ok(c) => c,
    Err(_) => Config::default(),
});
