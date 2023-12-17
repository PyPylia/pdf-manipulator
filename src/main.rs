#![feature(extract_if, path_file_prefix)]
#![windows_subsystem = "windows"]

use converter::OfficeConverter;

mod combiner;
mod converter;
mod gui;

fn main() -> eyre::Result<()> {
    let converter = OfficeConverter::new()?;

    gui::create_gui(converter)
}
