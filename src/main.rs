use clap::Parser;
use regex::Regex;
use rust_xlsxwriter::{Format, Workbook, XlsxError};
use std::collections::{HashMap, VecDeque};
use std::fs;
use std::path::PathBuf;

struct Solicitud {
    nombre_del_estudiante: String,
    plan_de_estudios: String,
    numero_solicitud: String,
    fecha_de_solicitud: String,
    identificacion: String,
    motivos: String,
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
            4 => Some(self.solicitud.fecha_de_solicitud.clone()),
            5 => Some(self.solicitud.identificacion.clone()),
            6 => Some(self.solicitud.motivos.clone()),
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

fn read_and_extract_data(pdf_contents: &str) -> Result<HashMap<String, Vec<Solicitud>>, String> {
    let id_sections_re = Regex::new(r"ID\s*\|\s*[A-Z]+\s*ID\s*\|\s*[A-Z\s]+").unwrap();
    let sections: Vec<&str> = id_sections_re.split(pdf_contents).collect();

    // Split the text into chunks, each starting with "SOLICITUD ESTUDIANTE"
    let solicitud_seccion_re = Regex::new(r"\d+\.\s*\|\s*SOLICITUD ESTUDIANTE\s*").unwrap();
    let mut chunks: VecDeque<&str> = VecDeque::new();
    for section in sections {
        let m: Vec<&str> = solicitud_seccion_re.split(section).collect();
        chunks.extend(m);
    }
    chunks.pop_front();
    

    // Simplified regex for a single solicitud block
    let solicitud_regex = Regex::new(r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s*\d*\s*(\d{1,2}/\d{1,2}/\d{4}|\d{1,2}/\d{1,2}/\d{2})\s*motivos\s*(.*)")
    .map_err(|e| format!("Error compiling regex: {}", e))?;
    let solicitud_regex2 = Regex::new(r"(?s)nombre del estudiante\s*(.+?)\s*identificación\s*(\d+)\s*plan de estudios\s*(.+?)\s*número y fecha de la solicitud\s*([^ ]+)\s*\d*\s*(\d{1,2}/\d{1,2}/\d{4}|\d{1,2}/\d{1,2}/\d{2})\s*")
    .map_err(|e| format!("Error compiling regex: {}", e))?;

    let mut cancelaciones_extemporanea_asignaturas: Vec<Solicitud> = Vec::new();
    let mut cancelaciones_semestre: Vec<Solicitud> = Vec::new();
    let mut registro_trabajo_grado: Vec<Solicitud> = Vec::new();

    for chunk in chunks.iter() {
        if let Some(captures) = solicitud_regex.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let plan_de_estudios = captures
                .get(3)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let fecha_de_solicitud = captures.get(5).map_or("", |m| m.as_str()).trim().to_string();
            let motivos = captures
                .get(6)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();


            if numero_solicitud.starts_with("CEA") {
                cancelaciones_extemporanea_asignaturas.push(Solicitud {
                    nombre_del_estudiante,
                    plan_de_estudios,
                    numero_solicitud,
                    fecha_de_solicitud,
                    identificacion,
                    motivos,
                })
            } else if numero_solicitud.starts_with("CS") {
                cancelaciones_semestre.push(Solicitud {
                    nombre_del_estudiante,
                    plan_de_estudios,
                    numero_solicitud,
                    fecha_de_solicitud,
                    identificacion,
                    motivos,
                })
            } else if numero_solicitud.starts_with("CT") {
                println!("Eh?");
                println!("{}", chunk)
            } else {
                println!("Eh?");
                println!("{}", chunk)
            }
        } else if let Some(captures) = solicitud_regex2.captures(chunk) {
            let nombre_del_estudiante = captures
                .get(1)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let identificacion = captures
                .get(2)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let plan_de_estudios = captures
                .get(3)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let numero_solicitud = captures
                .get(4)
                .map_or("", |m| m.as_str())
                .trim()
                .to_string();
            let fecha_str = captures.get(5).map_or("", |m| m.as_str()).trim();
            let fecha_de_solicitud = fecha_str.to_string();
            let motivos = "".to_string();

            if numero_solicitud.starts_with("RTG") {
                registro_trabajo_grado.push(Solicitud {
                    nombre_del_estudiante,
                    plan_de_estudios,
                    numero_solicitud,
                    fecha_de_solicitud,
                    identificacion,
                    motivos,
                });
            }
        } else {
            eprintln!("Warning: Could not parse solicitud in chunk:\n{}", chunk);
            eprintln!("--------------------");
        }
    }

    let mut solicitudes: HashMap<String, Vec<Solicitud>> = HashMap::new();
    solicitudes.insert(
        "CANC ASIGNATURAS".to_string(),
        cancelaciones_extemporanea_asignaturas,
    );
    solicitudes.insert("CANC SEMESTRE".to_string(), cancelaciones_semestre);
    solicitudes.insert(
        "REGISTRO TRABAJO DE GRADO".to_string(),
        registro_trabajo_grado,
    );

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
        if sheet_name.starts_with('R') {
        } else {
            worksheet.write_string_with_format(0, 5, "Motivos", &bold_format)?;
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

fn process_pdf(
    pdf_path: PathBuf,
    output_dir: PathBuf,
) -> Result<HashMap<String, Vec<Solicitud>>, String> {
    // Create the output directory if it doesn't exist
    if !output_dir.exists() {
        fs::create_dir_all(&output_dir)
            .map_err(|e| format!("Failed to create output directory: {}", e))?;
    }

    // Generate output paths
    let file_stem = pdf_path.file_stem().unwrap().to_string_lossy().to_string();
    let excel_path = output_dir.join(format!("{}.xlsx", file_stem));
    let bytes = std::fs::read(pdf_path).expect("Error reading PDF file crate");
    let out =
        pdf_extract::extract_text_from_mem(&bytes).expect("Error extracting text from PDF crate");
    // Check if pdftotext is available
    //if !is_pdftotext_available(&pdftotext_path) {
    //    return Err("pdftotext is not available. Please install it or provide a valid path.".to_string());
    //}

    // Convert PDF to text
    //convert_pdf_to_txt(&pdf_path, &txt_path, &pdftotext_path)?;

    // Extract data from text
    let data = read_and_extract_data(&out)?;

    // Write data to Excel
    write_data_to_excel(&data, &excel_path)
        .map_err(|e| format!("Failed to write Excel file: {}", e))?;

    Ok(data)
}

#[derive(Parser, Debug)]
#[command(
    author = "Jorge A. VM",
    version = "0.2.0",
    about = "PDF to Text Converter"
)]
struct Cli {
    #[arg(value_name = "PDF_PATH")]
    pdf_path: PathBuf,

    #[arg(short, long, default_value = ".", value_name = "OUTPUT_DIR")]
    output_dir: PathBuf,
}

fn main() {
    let cli = Cli::parse();

    let result = process_pdf(cli.pdf_path, cli.output_dir);

    match result {
        Ok(_) => println!("PDF processed successfully."),
        Err(e) => eprintln!("Error processing PDF: {}", e),
    }
}
