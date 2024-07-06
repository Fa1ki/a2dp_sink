use std::{ cell::RefCell, error::Error, io::{ stdin, stdout, Write }, rc::Rc };

use windows::{
    runtime::HSTRING,
    Devices::Enumeration::DeviceInformation,
    Foundation::TypedEventHandler,
    Media::Audio::{
        AudioPlaybackConnection,
        AudioPlaybackConnectionOpenResultStatus,
        AudioPlaybackConnectionState,
    },
};

fn main() {
    match run() {
        Ok(_) => (),
        Err(err) => println!("Error: {}", err),
    }
}

fn run() -> Result<(), Box<dyn Error>> {
    let selector = AudioPlaybackConnection::GetDeviceSelector()?;
    let device_watcher = DeviceInformation::CreateWatcherAqsFilter(selector)?;
    let device_list: Rc<RefCell<Vec<DeviceInformation>>> = Rc::new(RefCell::new(Vec::new()));

    let device_list_clone = device_list.clone();
    device_watcher
        .Added(
            TypedEventHandler::new(move |_sender, args: &Option<DeviceInformation>| {
                if let Some(device_info) = args.as_ref() {
                    println!(
                        "[DeviceWatcher] Added: \u{001B}[36m{}\u{001B}[0m",
                        device_info.Name().unwrap()
                    );
                    device_list_clone.borrow_mut().push(device_info.clone());
                }
                Ok(())
            })
        )
        .unwrap();

    device_watcher.Start()?;
    println!("Scanning for devices. Press enter to stop scanning and select a device.");
    let mut input = String::new();
    stdin().read_line(&mut input)?;

    device_watcher.Stop()?;

    let devices = device_list.borrow();
    if devices.is_empty() {
        println!("No devices found.");
        return Ok(());
    }

    println!("Select a device:");
    for (index, device) in devices.iter().enumerate() {
        println!("{}: \u{001B}[36m{}\u{001B}[0m", index, device.Name().unwrap());
    }

    let selected_index = loop {
        print!("Enter device number: ");
        stdout().flush().unwrap();
        input.clear();
        stdin().read_line(&mut input)?;

        match input.trim().parse::<usize>() {
            Ok(index) if index < devices.len() => {
                break index;
            }
            _ => println!("Invalid selection. Please try again."),
        }
    };

    let selected_device = &devices[selected_index];
    let _connection = connect(selected_device.Id().unwrap())?;
    println!("Connected to device: {}", selected_device.Name().unwrap());

    println!("Waiting for connection. Press enter to exit.");
    stdin().read_line(&mut input)?;
    Ok(())
}

fn format_state(state: AudioPlaybackConnectionState) -> String {
    match state {
        AudioPlaybackConnectionState::Opened => String::from("Opened"),
        AudioPlaybackConnectionState::Closed => String::from("Closed"),
        x => format!("{:?}", x),
    }
}

fn format_status(status: AudioPlaybackConnectionOpenResultStatus) -> String {
    match status {
        AudioPlaybackConnectionOpenResultStatus::Success => String::from("Success"),
        AudioPlaybackConnectionOpenResultStatus::DeniedBySystem => String::from("DeniedBySystem"),
        AudioPlaybackConnectionOpenResultStatus::RequestTimedOut => String::from("RequestTimedOut"),
        AudioPlaybackConnectionOpenResultStatus::UnknownFailure => String::from("UnknownFailure"),
        x => format!("{:?}", x),
    }
}

fn connect(device_id: HSTRING) -> Result<AudioPlaybackConnection, Box<dyn Error>> {
    let connection = AudioPlaybackConnection::TryCreateFromId(device_id)?;
    connection.StateChanged(
        TypedEventHandler::new(|sender: &Option<AudioPlaybackConnection>, _| {
            if let Some(connection) = sender.as_ref() {
                println!(
                    "[AudioPlaybackConnection] OnStateChanged: {}",
                    format_state(connection.State().unwrap())
                );
            }
            Ok(())
        })
    )?;
    connection.Start()?;
    let result = connection.Open()?;
    println!("[AudioPlaybackConnection] Open: {}", format_status(result.Status()?));
    Ok(connection)
}
