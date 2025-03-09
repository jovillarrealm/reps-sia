use chrono::{NaiveDate, ParseError};
use clap::Parser;
use eframe::egui;
use eframe::egui::IconData;
use image::ImageReader;
use once_cell::sync::Lazy;
use regex::Regex;
use rfd::FileDialog;
use rust_xlsxwriter::{Format, Workbook, XlsxError};
use std::collections::{HashMap, VecDeque};
use std::io::Cursor;
use std::path::PathBuf;

#[derive(Debug, Default)]
struct Solicitud {
    nombre_del_estudiante: String,
    plan_de_estudios: String,
    numero_solicitud: String,
    fecha_de_solicitud: NaiveDate,
    identificacion: usize,
    motivos: Option<String>,
    adjuntos: Option<usize>,
    materias: Option<String>,
    periodo: Option<String>,
}

struct SolicitudIterator<'a> {
    solicitud: &'a Solicitud,
    index: usize,
}

impl Iterator for SolicitudIterator<'_> {
    type Item = String;

    fn next(&mut self) -> Option<Self::Item> {
        self.index += 1;
        match self.index {
            1 => Some(self.solicitud.nombre_del_estudiante.clone()),
            2 => Some(self.solicitud.plan_de_estudios.clone()),
            3 => Some(self.solicitud.numero_solicitud.clone()),
            4 => Some(
                self.solicitud
                    .fecha_de_solicitud
                    .format("%d/%m/%Y")
                    .to_string(),
            ),
            5 => Some(self.solicitud.identificacion.to_string()),
            6 => match &self.solicitud.adjuntos {
                Some(anexos) => Some(anexos.to_string()),
                None => self.next(),
            },
            7 => match &self.solicitud.materias {
                Some(materias) => Some(materias.to_string()),
                None => self.next(),
            },
            8 => match &self.solicitud.periodo {
                Some(periodo) => Some(periodo.to_string()),
                None => self.next(),
            },
            9 => match &self.solicitud.motivos {
                Some(motivos) => Some(motivos.to_string()),
                None => self.next(),
            },
            _ => None,
        }
    }
}

impl Solicitud {
    fn iter(&self) -> SolicitudIterator {
        SolicitudIterator {
            solicitud: self,
            index: 0,
        }
    }
}

fn parse_date(date_str: &str) -> Result<NaiveDate, ParseError> {
    // Attempt to parse with both formats
    NaiveDate::parse_from_str(date_str, "%d/%m/%Y")
        .or_else(|_| NaiveDate::parse_from_str(date_str, "%d/%m/%y"))
}

const ANOTACIONES_SEP: &str = "ID |ESPACIO PARA ANOTACIONES";
const _HALF_SECTION_STR_SEP: &str = r"ID\s+\|[A-ZÁÉÍÓÚÜ\s]+";
const SOLICITUD_ESTUDIANTE_SEP: &str = r"\d+\.\s*\|\s*SOLICITUD ESTUDIANTE\s*";
const SECTION_STR: &str = concat!(r"ID\s+\|[A-ZÁÉÍÓÚÜ\s]+", "\n\n", r"ID\s+\|[A-ZÁÉÍÓÚÜ\s]+");

static ID_SECTIONS_RE: Lazy<Regex> = Lazy::new(|| Regex::new(SECTION_STR).unwrap());
static SOLICITUD_SECCION_RE: Lazy<Regex> =
    Lazy::new(|| Regex::new(SOLICITUD_ESTUDIANTE_SEP).unwrap());

const DOC_ANEX_DOC: &str = r"(\s*documento\s+anexo\s+Documento\s*)";
const ANOTACIONES: &str = "ANOTACIONES";

const CEA: &str = "CANCELACIÓN EXTEMP. ASIGNATURAS";
const RS_RE_CEA: &str = concat!(
    r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+\s*\d*)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s+(\d\s*\d/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s*",
    r"(?:motivos\s*(.*))(?:anexar otros documentos físicos\s*(.*))",
    r"(?:materias relacionadas a la solicitud\s*asignatura grp nombre(.*))",
);
static SOLICITUD_CEA: Lazy<Regex> = Lazy::new(|| {
    Regex::new(RS_RE_CEA)
    .map_err(|e| format!("Error compiling regex: {}", e))
    .unwrap()
});

const CS: &str = "CANCELACIÓN SEMESTRE";
const RS_RE_CS: &str = concat!(
    r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+\s*\d*)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s+(\d\s*\d/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s*",
    r"(?:motivos\s*(.*))(?:anexar otros documentos físicos\s*(.*))",
    r"(?:periodo para el que solicita cancelación de semestre\s*(.*))",
);
static SOLICITUD_CS: Lazy<Regex> = Lazy::new(|| {
    Regex::new(RS_RE_CS)
    .map_err(|e| format!("Error compiling regex: {}", e))
    .unwrap()
});


const ACM: &str = "AUTORIZACIÓN CARGA MÍNIMA";
const RS_RE_ACM: &str = concat!(
    r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+\s*\d*)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s+(\d\s*\d/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s*",
    r"(?:motivos\s*(.*))(?:anexar otros documentos físicos\s*(.*))",
    r"(?:periodo para el que solicita carga mínima\s*(.*))",
);
static SOLICITUD_ACM: Lazy<Regex> = Lazy::new(|| {
    Regex::new(RS_RE_ACM)
        .map_err(|e| format!("Error compiling regex: {}", e))
        .unwrap()
    });

    
const RTG: &str = "REGISTRO TRABAJO GRADO";
const RS_RE_RTG: &str = r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+\s*\d*)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s+(\d\s*\d/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s*(?:anexar otros documentos físicos\s*(.*))";
static SOLICITUD_RTG: Lazy<Regex> = Lazy::new(|| {
    Regex::new(RS_RE_RTG)
        .map_err(|e| format!("Error compiling regex: {}", e))
        .unwrap()
});

static DOC_ANEX_DOC_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(DOC_ANEX_DOC)
        .map_err(|e| format!("Error compiling regex: {}", e))
        .unwrap()
});

fn read_and_extract_data(pdf_contents: &str) -> Result<HashMap<String, Vec<Solicitud>>, String> {
    // Split anotaciones from the rest of sections

    let pdf_contents: Vec<&str> = pdf_contents.split(ANOTACIONES_SEP).collect();
    let anotaciones = pdf_contents[1];
    let pdf_contents = pdf_contents[0];

    // Separate pdf by ID | sections
    let sections: Vec<&str> = ID_SECTIONS_RE.split(pdf_contents).collect();

    // Split each section into chunks, each starting with "SOLICITUD ESTUDIANTE"
    let mut chunks: VecDeque<&str> = VecDeque::new();
    for section in sections {
        let m: Vec<&str> = SOLICITUD_SECCION_RE.split(section).collect();
        chunks.extend(m);
    }
    // First chunk just has information about this pdf, we are ignoring that
    chunks.pop_front();
    // We now have each
    // Clean that shit.
    let chunks = chunks
        .into_iter()
        .filter(|s| !s.is_empty())
        .collect::<Vec<&str>>();

    // Simplified regex for a single solicitud block
    let _rs_re = r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+\s*\d*)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s+(\d\s*\d/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s*";

    let mut cancelaciones_extemporanea_asignaturas: Vec<Solicitud> = Vec::new();
    let mut cancelacion_extemporanea_asignaturas_posgrado: Vec<Solicitud> = Vec::new();
    let mut cancelaciones_semestre: Vec<Solicitud> = Vec::new();
    let mut registro_trabajo_grado: Vec<Solicitud> = Vec::new();
    let mut autorizacion_menor_carga_minima: Vec<Solicitud> = Vec::new();

    for chunk in chunks.iter() {
        if let Some(captures) = SOLICITUD_CEA.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("")
                .parse()
                .expect("Error parsing identificacion");
            let plan_de_estudios = captures.get(3).map_or("", |m| m.as_str()).trim().into();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let n_sol = numero_solicitud.clone();
            let fecha_de_solicitud_str = captures
                .get(5)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("");
            let fecha_de_solicitud = match parse_date(&fecha_de_solicitud_str) {
                Ok(date) => date,
                Err(e) => {
                    eprintln!("Error parsing date '{}': {}", fecha_de_solicitud_str, e);
                    return Err(format!(
                        "Error parsing date '{}': {}",
                        fecha_de_solicitud_str, e
                    ));
                }
            };

            let periodo: Option<String> = None;
            let motivos = Some(
                captures
                    .get(6)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            // solicitud empieza con CEA so far
            let _anexos = Some(
                captures
                    .get(7)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            let last_field_capture = captures.get(8).map_or("", |m| m.as_str()).trim();
            let adjuntos = Some(DOC_ANEX_DOC_RE.find_iter(last_field_capture).count());
            let second = DOC_ANEX_DOC_RE.replace(last_field_capture, "");
            let materias = Some(second.to_string());
            let solicitud = Solicitud {
                nombre_del_estudiante,
                plan_de_estudios,
                numero_solicitud,
                fecha_de_solicitud,
                identificacion,
                motivos,
                adjuntos,
                materias,
                periodo,
            };
            if n_sol.starts_with("CEAP") {
                cancelacion_extemporanea_asignaturas_posgrado.push(solicitud)
            } else if n_sol.starts_with("CEA") {
                cancelaciones_extemporanea_asignaturas.push(solicitud)
            } else {
                println!("Warning: Eh? QUE ES ESTO?????");
                println!("{}", chunk)
            }
        } else if let Some(captures) = SOLICITUD_CS.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("")
                .parse()
                .expect("Error parsing identificacion");
            let plan_de_estudios = captures.get(3).map_or("", |m| m.as_str()).trim().into();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let n_sol = numero_solicitud.clone();
            let fecha_de_solicitud_str = captures
                .get(5)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("");
            let fecha_de_solicitud = match parse_date(&fecha_de_solicitud_str) {
                Ok(date) => date,
                Err(e) => {
                    eprintln!("Error parsing date '{}': {}", fecha_de_solicitud_str, e);
                    return Err(format!(
                        "Error parsing date '{}': {}",
                        fecha_de_solicitud_str, e
                    ));
                }
            };

            let motivos = Some(
                captures
                    .get(6)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            let _anexos = Some(
                captures
                    .get(7)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            let materias = None;
            let last_field_capture = captures.get(8).map_or("", |m| m.as_str()).trim();
            let adjuntos = Some(DOC_ANEX_DOC_RE.find_iter(last_field_capture).count());
            let second = DOC_ANEX_DOC_RE.replace(last_field_capture, "");
            let periodo = Some(second.to_string());

            let solicitud = Solicitud {
                nombre_del_estudiante,
                plan_de_estudios,
                numero_solicitud,
                fecha_de_solicitud,
                identificacion,
                motivos,
                adjuntos,
                materias,
                periodo,
            };
            if n_sol.starts_with("CS") {
                cancelaciones_semestre.push(solicitud);
            } else if n_sol.starts_with("ACM")  {
                autorizacion_menor_carga_minima.push(solicitud);
                
            } else {
                dbg!(captures.get(0));
            }
        } else if let Some(captures) = SOLICITUD_ACM.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("")
                .parse()
                .expect("Error parsing identificacion");
            let plan_de_estudios = captures.get(3).map_or("", |m| m.as_str()).trim().into();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let n_sol = numero_solicitud.clone();
            let fecha_de_solicitud_str = captures
                .get(5)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("");
            let fecha_de_solicitud = match parse_date(&fecha_de_solicitud_str) {
                Ok(date) => date,
                Err(e) => {
                    eprintln!("Error parsing date '{}': {}", fecha_de_solicitud_str, e);
                    return Err(format!(
                        "Error parsing date '{}': {}",
                        fecha_de_solicitud_str, e
                    ));
                }
            };

            let motivos = Some(
                captures
                    .get(6)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            let _anexos = Some(
                captures
                    .get(7)
                    .map_or("", |m| m.as_str())
                    .trim()
                    .to_string(),
            );
            let materias = None;
            let last_field_capture = captures.get(8).map_or("", |m| m.as_str()).trim();
            let adjuntos = Some(DOC_ANEX_DOC_RE.find_iter(last_field_capture).count());
            let second = DOC_ANEX_DOC_RE.replace(last_field_capture, "");
            let periodo = Some(second.to_string());

            let solicitud = Solicitud {
                nombre_del_estudiante,
                plan_de_estudios,
                numero_solicitud,
                fecha_de_solicitud,
                identificacion,
                motivos,
                adjuntos,
                materias,
                periodo,
            };

            if n_sol.starts_with("ACM") {
                autorizacion_menor_carga_minima.push(solicitud);
            } else {
                dbg!(captures.get(0));
            }
        } else if let Some(captures) = SOLICITUD_RTG.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("")
                .parse()
                .expect("Error parsing identificacion");
            let plan_de_estudios = captures.get(3).map_or("", |m| m.as_str()).trim().into();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let n_sol = numero_solicitud.clone();
            let fecha_de_solicitud_str = captures
                .get(5)
                .map_or("", |m| m.as_str())
                .split_whitespace()
                .collect::<Vec<&str>>()
                .join("");
            let fecha_de_solicitud = match parse_date(&fecha_de_solicitud_str) {
                Ok(date) => date,
                Err(e) => {
                    eprintln!("Error parsing date '{}': {}", fecha_de_solicitud_str, e);
                    return Err(format!(
                        "Error parsing date '{}': {}",
                        fecha_de_solicitud_str, e
                    ));
                }
            };
            let motivos = None;
            let materias = None;
            
            let last_field_capture = captures.get(6).map_or("", |m| m.as_str());
            
            let adjuntos = Some(DOC_ANEX_DOC_RE.find_iter(last_field_capture).count());

            let _second = DOC_ANEX_DOC_RE.replace(last_field_capture, "");
            let periodo = None;
            let solicitud = Solicitud {
                nombre_del_estudiante,
                plan_de_estudios,
                numero_solicitud,
                fecha_de_solicitud,
                identificacion,
                motivos,
                adjuntos,
                materias,
                periodo,
            };
            if n_sol.starts_with("RTG") {
                registro_trabajo_grado.push(solicitud);
            } else {
                eprintln!("Warning: Could not parse solicitud in chunk:\n{}", chunk);
                eprintln!("--------------------");
            }

        } else {
            eprintln!("Warning: Could not parse solicitud in chunk:\n{}", chunk);
            eprintln!("--------------------");
        }
    }

    let mut solicitudes: HashMap<String, Vec<Solicitud>> = HashMap::new();

    if !cancelacion_extemporanea_asignaturas_posgrado.is_empty() {
        //solicitudes.insert(
        //    "CANCELACIÓN EXTEMP. ASIGN. POS".to_string(),
        //    cancelacion_extemporanea_asignaturas_posgrado,
        //);
        cancelaciones_extemporanea_asignaturas
            .extend(cancelacion_extemporanea_asignaturas_posgrado);
    }
    if !cancelaciones_extemporanea_asignaturas.is_empty() {
        solicitudes.insert(CEA.to_string(), cancelaciones_extemporanea_asignaturas);
    }

    if !cancelaciones_semestre.is_empty() {
        solicitudes.insert(CS.to_string(), cancelaciones_semestre);
    }
    if !registro_trabajo_grado.is_empty() {
        solicitudes.insert(RTG.to_string(), registro_trabajo_grado);
    }
    if !autorizacion_menor_carga_minima.is_empty() {
        solicitudes.insert(ACM.to_string(), autorizacion_menor_carga_minima);
    }
    if !anotaciones.is_empty() {
        let bs = Solicitud {
            motivos: Some(anotaciones.to_string()),
            ..Default::default()
        };
        solicitudes.insert(ANOTACIONES.to_string(), vec![bs]);
    }

    Ok(solicitudes)
}

fn write_data_to_excel(
    data: &HashMap<String, Vec<Solicitud>>,
    excel_path: &PathBuf,
) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    for (sheet_name, sheet_data) in data {
        let worksheet = workbook.add_worksheet().set_name(sheet_name)?;

        // Add a bold format for the headers.
        let bold_format = Format::new().set_bold();

        // Write the headers.
        worksheet.write_string_with_format(0, 0, "Nombre del estudiante", &bold_format)?;
        worksheet.write_string_with_format(0, 1, "Plan de estudios", &bold_format)?;
        worksheet.write_string_with_format(0, 2, "Número de solicitud", &bold_format)?;
        worksheet.write_string_with_format(0, 3, "Fecha de solicitud", &bold_format)?;
        worksheet.write_string_with_format(0, 4, "Identificación", &bold_format)?;
        let mut j = 5;
        if sheet_name.starts_with(CEA) {
            worksheet.write_string_with_format(0, j, "Adjuntos", &bold_format)?;
            j += 1;
            worksheet.write_string_with_format(0, j, "Materias", &bold_format)?;
            j += 1;
        } else if sheet_name.starts_with(CS) || sheet_name.starts_with(ACM) {
            worksheet.write_string_with_format(0, j, "Adjuntos", &bold_format)?;
            j += 1;
            worksheet.write_string_with_format(0, j, "Periodo", &bold_format)?;
            j += 1;
        }
        if !sheet_name.starts_with(RTG) {
            worksheet.write_string_with_format(0, j, "Motivos", &bold_format)?;
        }

        // Write data rows
        for (row, sol) in sheet_data.iter().enumerate() {
            for (col, field) in sol.iter().enumerate() {
                worksheet.write_string(row as u32 + 1, col as u16, field)?;
            }
        }
    }

    workbook.save(excel_path)?;

    Ok(())
}

fn process_pdf(pdf_path: PathBuf) -> Result<HashMap<String, Vec<Solicitud>>, String> {
    // Generate output paths
    let file_stem = pdf_path.file_stem().unwrap().to_string_lossy().to_string();
    let excel_path = pdf_path
        .parent()
        .unwrap()
        .join(format!("{}.xlsx", file_stem));
    let excel_name = excel_path
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();

    if let Some(output_path) = FileDialog::new()
        .add_filter("Excel", &["xlsx"])
        .set_file_name(excel_name)
        .save_file()
    {
        // Your existing processing logic should be integrated here
        let bytes = std::fs::read(pdf_path).expect("Error reading PDF file crate");
        let out = pdf_extract::extract_text_from_mem(&bytes)
            .expect("Error extracting text from PDF crate");

        // Extract data from text
        let data = read_and_extract_data(&out)?;

        // Write data to Excel
        write_data_to_excel(&data, &output_path)
            .map_err(|e| format!("Failed to write Excel file: {}", e))?;
        Ok(data)
    } else {
        Err("No output file selected".to_string())
    }
}

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

#[derive(Parser, Debug)]
#[command(
    author = "Jorge A. VM",
    about = "PDF to Text Converter"
)]
struct Cli {
    #[arg(value_name = "PDF_PATH")]
    pdf_path: PathBuf,

    #[arg(short, long, default_value = ".", value_name = "OUTPUT_DIR")]
    output_dir: PathBuf,
}

struct PdfProcessorApp {
    pdf_path: Option<PathBuf>,
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

            if ui.button("Seleccione PDF").clicked() {
                if let Some(path) = FileDialog::new().add_filter("PDF", &["pdf"]).pick_file() {
                    self.pdf_path = Some(path);
                }
            }

            if let Some(path) = &self.pdf_path {
                ui.label(format!("PDF Seleccionado: {}", path.display()));
            }

            if ui.button("Procesar PDF").clicked() {
                if let Some(path) = &self.pdf_path {
                    match process_pdf(path.clone()) {
                        Ok(_) => self.status = format!("{} procesado\n\nValida que el excel no contenga errores, esto aún es experimental.\nEspero sea de su agrado.\n\t-JAVM", path.clone().file_name().expect("pdf path").to_string_lossy()),
                        Err(e) => self.status = format!("Error: {}", e),
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
