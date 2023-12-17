use reqwest::blocking::Client;
use std::{fs::File, io::Write, path::Path};
use winres::WindowsResource;

const OFFICETOPDF_URL: &str = "https://api.github.com/repos/cognidox/OfficeToPDF/releases/latest";

fn main() {
    let mut res = WindowsResource::new();
    res.set_icon("AppIcon.ico");
    res.compile().expect("Unable to set app icon.");

    if !(Path::new("target/OfficeToPDF.exe").exists() && cfg!(debug_assertions)) {
        download_officetopdf()
    }
}

fn download_officetopdf() {
    let client = Client::new();

    let query_request = client
        .get(OFFICETOPDF_URL)
        .header("Accept", "application/vnd.github+json")
        .header("X-GitHub-Api-Version", "2022-11-28")
        .header("User-Agent", "reqwest/0.11.22");
    let query_response = query_request.send().expect("Query request failed.");
    let query_body = query_response
        .text()
        .expect("Could not get query response body.");
    let query_json = json::parse(&query_body).expect("Could not parse query response json.");
    let download_url = query_json["assets"]
        .members()
        .find(|obj| obj["name"] == "OfficeToPDF.exe")
        .expect("Could not find valid OfficeToPDF executable.")["browser_download_url"]
        .as_str()
        .expect("Could not get OfficeToPDF download URL.");

    let executable_response = client
        .get(download_url)
        .send()
        .expect("Could not download OfficeToPDF.");
    let executable_data = executable_response
        .bytes()
        .expect("Could not get executable data for OfficeToPDF.");

    let mut executable_file = File::create("target/OfficeToPDF.exe")
        .expect("Could not open OfficeToPDF executable file.");
    executable_file
        .write_all(&executable_data)
        .expect("Could not write OfficeToPDF executable data.");
}
