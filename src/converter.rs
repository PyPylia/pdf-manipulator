use std::{env, fs, os::windows::process::CommandExt, path::PathBuf, process::Command};
use uuid::Uuid;

static OFFICE_TO_PDF: &[u8] = include_bytes!("../target/OfficeToPDF.exe");

pub struct OfficeConverter {
    executable_path: PathBuf,
}

impl OfficeConverter {
    pub fn new() -> eyre::Result<OfficeConverter> {
        let app_data = PathBuf::from(env::var("APPDATA")?);

        let mut app_folder = app_data;
        app_folder.push("pdf-manipulator");
        if !app_folder.exists() {
            fs::create_dir(&app_folder)?
        }

        let mut office_to_pdf = app_folder;
        office_to_pdf.push("OfficeToPDF.exe");
        if !office_to_pdf.exists() {
            fs::write(&office_to_pdf, OFFICE_TO_PDF)?;
        }

        Ok(OfficeConverter {
            executable_path: office_to_pdf,
        })
    }

    pub fn convert_files(&self, files: &[PathBuf]) -> Vec<PathBuf> {
        let mut children = vec![];

        for file in files {
            let output_id = Uuid::new_v4();
            let output_file = env::temp_dir().join(output_id.simple().to_string() + ".pdf");

            let child = Command::new(&self.executable_path)
                .arg(file)
                .arg(&output_file)
                .creation_flags(0x08000000) // CREATE_NO_WINDOW
                .spawn()
                .expect("Couldn't run OfficeToPDF.");

            children.push((output_file, child));
        }

        children
            .into_iter()
            .filter_map(|(output_file, mut handle)| {
                if handle
                    .wait()
                    .expect("Unexpected error while waiting on child process.")
                    .success()
                {
                    Some(output_file)
                } else {
                    None
                }
            })
            .collect()
    }
}
