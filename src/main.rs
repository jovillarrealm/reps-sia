//use clap::Parser;
use eframe::egui;
use eframe::egui::IconData;
use image::ImageReader;
use rfd::FileDialog;
use std::io::Cursor;
use std::path::PathBuf;
mod pdf_handling;


fn load_icon() -> Option<IconData> {
    let icon_bytes = include_bytes!("favicon.png"); // Replace with your favicon file
    let image = ImageReader::new(Cursor::new(icon_bytes))
        .with_guessed_format()
        .ok()?
        .decode()
        .ok()?
        .into_rgba8();
    let (width, height) = image.dimensions();
    Some(IconData {
        rgba: image.into_raw(),
        width,
        height,
    })
}

/*
#[derive(Parser, Debug)]
#[command(author = "Jorge A. VM", about = "PDF to Text Converter")]
struct Cli {
    #[arg(value_name = "PDF_PATH")]
    pdf_path: PathBuf,
    
    #[arg(short, long, default_value = ".", value_name = "OUTPUT_DIR")]
    output_dir: PathBuf,
}
*/

struct PdfProcessorApp {
    pdf_path: Option<Vec<PathBuf>>,
    status: String,
}

impl Default for PdfProcessorApp {
    fn default() -> Self {
        Self {
            pdf_path: None,
            status: "Seleccione un archivo de PDF".to_string(),
        }
    }
}
impl eframe::App for PdfProcessorApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Procesador de Reportes de Agenda del SIA");

            if ui.button("Seleccione PDF(s)").clicked() {
                if let Some(path) = FileDialog::new().add_filter("PDF", &["pdf"]).pick_files() {
                    self.pdf_path = Some(path);
                }
            }

            if let Some(paths) = &self.pdf_path {
                ui.label(format!("PDF Seleccionado: {}", paths
                .iter()
                .map(|path| path.display().to_string())
                .collect::<Vec<String>>()
                .join("\n")));
            }

            if ui.button("Procesar PDF").clicked() {
                if let Some(paths) = &self.pdf_path {
                    for pdf_path in paths{
                        match pdf_handling::process_pdf(pdf_path.clone()) {
                            Ok(_) => self.status = format!("{} procesado\n\n Por favor valida que el excel no contenga errores, este software aún es experimental.\n\nEn caso de haber solicitudes sin manejar, se intentan escribir en un archivo txt\n\t-JAVM", pdf_path.clone().file_name().expect("pdf path").to_string_lossy()),
                            Err(e) => self.status = format!("Error: {e}"),
                        }
                    }
                }
            }

            ui.label(&self.status);
        });
    }
}

fn main() -> Result<(), eframe::Error> {
    let icon = load_icon().unwrap();
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_icon(icon),
        ..Default::default()
    }; // Start app with icon

    eframe::run_native(
        "Procesador de PDF de Reporte de Agenda a Excel",
        options,
        Box::new(|_cc| Ok(Box::new(PdfProcessorApp::default()))),
    )?;
    Ok(())
}
