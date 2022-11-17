use chrono::{DateTime, Datelike, Duration, Local, TimeZone, Timelike};
use log::{error, info, warn, LevelFilter};
use regex::Regex;
use std::collections::HashMap;
use std::error::Error;
use std::thread;

use lazy_static::lazy_static;

use wmi::{COMLibrary, Variant, WMIConnection};

fn get_serial() -> Result<String, Box<dyn Error>> {
    lazy_static! {
        static ref REGEX_SERIAL_PORT: Regex = Regex::new(r"USB-SERIAL CH340 \((COM\d+)\)").unwrap();
    }
    let query_string =  "SELECT Caption FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"";

    let conn = WMIConnection::new(COMLibrary::new()?)?;
    let results: Vec<HashMap<String, Variant>> = conn.raw_query(query_string)?;
    let mut result_ports: Vec<String> = Vec::new();
    for result in results {
        if let Some(Variant::String(caption)) = result.get("Caption") {
            if let Some(cap) = REGEX_SERIAL_PORT.captures(caption) {
                let port_number = cap.get(1).unwrap().as_str();
                result_ports.push(port_number.into());
            }
        }
    }

    if result_ports.is_empty() {
        return Err("No serial ports found".into());
    }

    if result_ports.len() > 1 {
        warn!(
            "Multiple serial ports found, using first one: {}",
            result_ports[0]
        );
    }

    Ok(result_ports.into_iter().next().unwrap())
}

fn time_trunc_second(time: &DateTime<Local>) -> DateTime<Local> {
    Local
        .with_ymd_and_hms(
            time.year(),
            time.month(),
            time.day(),
            time.hour(),
            time.minute(),
            time.second(),
        )
        .unwrap()
}

fn construct_data_buf(time: impl Timelike) -> [u8; 6] {
    let seconds = ((time.hour() * 60) + time.minute()) * 60 + time.second();
    let mut result = *b"Sb\x00\x00\x00\x00";
    result[5] = ((seconds & 0x7f) | 0x80) as u8;
    result[4] = (((seconds >> 7) & 0x7f) | 0x80) as u8;
    result[3] = (((seconds >> 14) & 0x7f) | 0x80) as u8;
    result[2] = (((seconds >> 21) & 0x7f) | 0x80) as u8;
    result
}

fn main() {
    env_logger::builder()
        .filter_level(LevelFilter::Trace)
        .init();
    if let Err(e) = inner_main() {
        error!("main error: {}", e);
    }
}

fn inner_main() -> Result<(), Box<dyn Error>> {
    let serial_port_num = get_serial()?;
    info!("Serial port number: {}", serial_port_num);
    let mut serial = serialport::new(serial_port_num, 115200).open()?;
    let now = Local::now();
    let mut next = now;
    let (next_sync_time, dist) = loop {
        next = next + Duration::seconds(1);
        let next_sync_time = time_trunc_second(&next);
        let dist = next_sync_time - now;
        if dist > Duration::microseconds(100) {
            break (next_sync_time, dist);
        }
    };

    let buf = construct_data_buf(next_sync_time);
    serial.write(&buf)?;

    let sleep_duration = next_sync_time - Local::now();
    if sleep_duration < Duration::zero() {
        error!("Failed to finish operation within {:?}", dist);
        return Err("Failed to finish operation.".into());
    }
    let sleep_duration = sleep_duration.to_std()?;
    thread::sleep(sleep_duration);

    serial.write(b"c")?;

    info!("Sync finished to time {}", next_sync_time);

    Ok(())
}
