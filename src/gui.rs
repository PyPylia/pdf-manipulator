use crate::{combiner::combine_pdfs, converter::OfficeConverter};
use native_windows_derive::NwgUi;
use native_windows_gui::{
    modal_error_message, modal_info_message, stop_thread_dispatch, Button, EventData, FileDialog,
    Font, GridLayout, Icon, InsertListViewColumn, ListView, ListViewColumnFlags, NativeUi, Notice,
    ProgressBar, ProgressBarFlags, TextInput, Window,
};
#[allow(unused_imports)]
use native_windows_gui::{FileDialogAction, ListViewExFlags, ListViewStyle};
use std::{
    cell::RefCell,
    env,
    ffi::OsStr,
    mem,
    path::PathBuf,
    thread::{self, JoinHandle},
};

#[derive(NwgUi)]
pub struct MainGui {
    #[nwg_resource(source_bin: Some(include_bytes!("../AppIcon.ico")))]
    app_icon: Icon,

    #[nwg_control(size: (960, 540), title: "PDF Manipulator", icon: Some(&data.app_icon))]
    #[nwg_events(OnWindowClose: [MainGui::exit], OnInit: [MainGui::initialise])]
    window: Window,

    #[nwg_layout(parent: window, max_row: Some(10), max_column: Some(7))]
    main_layout: GridLayout,

    #[nwg_resource(
        title: "Select Files",
        action: FileDialogAction::Open,
        multiselect: true,
        filters: "Any Document (*.*)",
    )]
    file_dialog: FileDialog,

    #[nwg_resource(
        title: "Select Output File",
        action: FileDialogAction::Save,
        filters: "Pdf Documents (*.pdf)",
    )]
    output_dialog: FileDialog,

    #[nwg_control(
        item_count: 0,
        list_style: ListViewStyle::Detailed,
        ex_flags: ListViewExFlags::GRID | ListViewExFlags::FULL_ROW_SELECT,
    )]
    #[nwg_layout_item(layout: main_layout, col: 0, row: 0, row_span: 8, col_span: 7)]
    #[nwg_events(OnListViewClick: [MainGui::files_clicked(SELF, EVT_DATA)])]
    file_list: ListView,

    #[nwg_control(text: "Add Files", focus: true)]
    #[nwg_layout_item(layout: main_layout, col: 0, row: 8)]
    #[nwg_events(OnButtonClick: [MainGui::add_files])]
    add_file_btn: Button,

    #[nwg_control(text: "Select Output")]
    #[nwg_layout_item(layout: main_layout, col: 1, row: 8)]
    #[nwg_events(OnButtonClick: [MainGui::set_output_file])]
    set_output_btn: Button,

    #[nwg_control(readonly: true)]
    #[nwg_layout_item(layout: main_layout, col: 2, row: 8, col_span: 4)]
    output_file: TextInput,

    #[nwg_control(text: "Process Files")]
    #[nwg_layout_item(layout: main_layout, col: 6, row: 8)]
    #[nwg_events(OnButtonClick: [MainGui::process_files])]
    process_files_btn: Button,

    #[nwg_control]
    #[nwg_layout_item(layout: main_layout, col: 0, row: 9, col_span: 7)]
    progress_bar: ProgressBar,

    #[nwg_control]
    #[nwg_events(OnNotice: [MainGui::processed_files])]
    processed_notice: Notice,

    backend: RefCell<BackendWrapper>,
    files: RefCell<Vec<PathBuf>>,
}

enum BackendWrapper {
    Running(JoinHandle<OfficeConverter>),
    Idle(OfficeConverter),
    Poisoned,
}

impl MainGui {
    fn initialise(&self) {
        if let Ok(user_profile) = env::var("USERPROFILE") {
            let mut documents = PathBuf::from(user_profile);
            documents.push("Documents");

            let documents_str = documents.to_string_lossy();

            self.file_dialog
                .set_default_folder(&documents_str)
                .expect("Couldn't set default folder for file dialog.");
            self.output_dialog
                .set_default_folder(&documents_str)
                .expect("Couldn't set default folder for output file dialog.")
        }

        self.file_list.insert_column(InsertListViewColumn {
            index: Some(0),
            fmt: Some(ListViewColumnFlags::LEFT),
            width: Some(700),
            text: Some("File name".into()),
        });
        self.file_list.insert_column(InsertListViewColumn {
            index: Some(1),
            fmt: Some(ListViewColumnFlags::LEFT),
            width: Some(100),
            text: Some("Document type".into()),
        });
        self.file_list.insert_column(InsertListViewColumn {
            index: Some(2),
            fmt: Some(ListViewColumnFlags::CENTER),
            width: Some(25),
            text: None,
        });
        self.file_list.set_headers_enabled(true);
    }

    fn add_files(&self) {
        if self.file_dialog.run(Some(&self.window)) {
            if let Ok(files) = self.file_dialog.get_selected_items() {
                for file in files {
                    let path = PathBuf::from(file);
                    if self.files.borrow().contains(&path) {
                        continue;
                    }

                    self.file_list.insert_items_row(
                        None,
                        &[
                            path.file_prefix()
                                .expect("No file prefix for file.")
                                .to_string_lossy()
                                .as_ref(),
                            path.extension()
                                .expect("No file extension for file.")
                                .to_string_lossy()
                                .as_ref(),
                            "\u{274C}",
                        ],
                    );

                    self.files.borrow_mut().push(path);
                }
            }
        }
    }

    fn set_output_file(&self) {
        if self.output_dialog.run(Some(&self.window)) {
            self.output_file.set_text("");
            if let Ok(output_file) = self.output_dialog.get_selected_item() {
                let mut text = output_file.to_string_lossy();
                if !text.ends_with(".pdf") {
                    text.to_mut().push_str(".pdf");
                }
                self.output_file.set_text(&text);
            }
        }
    }

    fn process_files(&self) {
        if self.files.borrow().len() == 0 {
            modal_error_message(
                &self.window,
                "Invalid input",
                "Please add files to process.",
            );
            return;
        }

        if self.output_file.text().is_empty() {
            modal_error_message(
                &self.window,
                "Invalid output",
                "Please select an output file.",
            );
            return;
        }

        let mut backend = self.backend.borrow_mut();
        if let BackendWrapper::Idle(converter) =
            mem::replace(&mut *backend, BackendWrapper::Poisoned)
        {
            self.process_files_btn.set_enabled(false);
            self.add_file_btn.set_enabled(false);
            self.set_output_btn.set_enabled(false);

            let finish_notice = self.processed_notice.sender();

            let mut files = self.files.borrow().clone();
            let output_file = PathBuf::from(self.output_file.text());
            let office_documents: Vec<PathBuf> = files
                .extract_if(|path| path.extension() != Some(OsStr::new("pdf")))
                .collect();

            self.progress_bar.set_marquee(true, 30);
            self.progress_bar.add_flags(ProgressBarFlags::MARQUEE);

            *backend = BackendWrapper::Running(thread::spawn(move || {
                let mut converted_files = converter.convert_files(office_documents.as_slice());

                files.append(&mut converted_files);
                combine_pdfs(files, &output_file);

                finish_notice.notice();
                converter
            }));
        }
    }

    fn processed_files(&self) {
        let mut backend = self.backend.borrow_mut();
        if let BackendWrapper::Running(handle) =
            mem::replace(&mut *backend, BackendWrapper::Poisoned)
        {
            *backend = BackendWrapper::Idle(
                handle
                    .join()
                    .expect("Unexpected error while waiting on backend thread."),
            );
        }

        self.progress_bar.set_marquee(false, 0);
        self.progress_bar.remove_flags(ProgressBarFlags::MARQUEE);

        self.process_files_btn.set_enabled(true);
        self.add_file_btn.set_enabled(true);
        self.set_output_btn.set_enabled(true);

        modal_info_message(
            &self.window,
            "Processing complete!",
            "Successfully processed files.",
        );
    }

    fn files_clicked(&self, event_data: &EventData) {
        let (row, column) = event_data.on_list_view_item_index();
        if column == 2 && self.file_list.has_item(row, column) {
            self.file_list.remove_item(row);
            self.files.borrow_mut().remove(row);
        }
    }

    fn exit(&self) {
        stop_thread_dispatch();
    }
}

pub fn create_gui(converter: OfficeConverter) -> eyre::Result<()> {
    native_windows_gui::init()?;
    Font::set_global_family("Segoe UI")?;

    let app = MainGui {
        app_icon: Icon::default(),
        window: Window::default(),
        main_layout: GridLayout::default(),
        file_dialog: FileDialog::default(),
        add_file_btn: Button::default(),
        file_list: ListView::default(),
        files: Default::default(),
        output_file: TextInput::default(),
        output_dialog: FileDialog::default(),
        set_output_btn: Button::default(),
        process_files_btn: Button::default(),
        progress_bar: ProgressBar::default(),
        processed_notice: Notice::default(),
        backend: RefCell::new(BackendWrapper::Idle(converter)),
    };
    let _ui = MainGui::build_ui(app)?;

    native_windows_gui::dispatch_thread_events();

    Ok(())
}
