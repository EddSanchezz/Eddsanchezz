from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pathlib import Path
from datetime import datetime
import shutil, os

TARGET = Path(r'C:\Users\metzu\Downloads\Hoja_de_Vida_Edward_Sanchez_Motta.docx')
TEMP_OUT = Path(r'C:\Users\metzu\AppData\Local\Temp\opencode\Hoja_de_Vida_Edward_Sanchez_Motta_updated.docx')
BACKUP = TARGET.with_name(TARGET.stem + '_original_backup' + TARGET.suffix)
NAVY, GRAY = '17365D', '4B5563'


def set_font(run, size=10, bold=None, color=None, italic=None, name='Arial'):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None:
        run.italic = italic
    return run


def hyperlink(paragraph, text, url, bold=False, size=9.2):
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement('w:hyperlink')
    link.set(qn('r:id'), rid)
    run = OxmlElement('w:r')
    props = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'), 'Arial')
    fonts.set(qn('w:hAnsi'), 'Arial')
    props.append(fonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(size * 2)))
    props.append(sz)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), NAVY)
    props.append(color)
    if bold:
        props.append(OxmlElement('w:b'))
    run.append(props)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    link.append(run)
    paragraph._p.append(link)


def bottom_border(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for key, value in [('val', 'single'), ('sz', '6'), ('space', '2'), ('color', 'B8C4D1')]:
        bottom.set(qn('w:' + key), value)
    borders.append(bottom)
    ppr.append(borders)


def keep(p, next_=False):
    p.paragraph_format.keep_together = True
    if next_:
        p.paragraph_format.keep_with_next = True


def section(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    keep(p, True)
    set_font(p.add_run(text.upper()), 10.2, True, NAVY)
    bottom_border(p)
    return p


def body(doc, text, size=9.4, after=1.3):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0
    keep(p)
    set_font(p.add_run(text), size)
    return p


def skill(doc, label, content):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0.6)
    p.paragraph_format.line_spacing = 1.0
    keep(p)
    set_font(p.add_run(label + ': '), 9.4, True)
    set_font(p.add_run(content), 9.4)
    return p


def project(doc, name, url, stack, description):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.13)
    p.paragraph_format.first_line_indent = Inches(-0.13)
    p.paragraph_format.space_after = Pt(1.4)
    p.paragraph_format.line_spacing = 1.0
    keep(p)
    set_font(p.add_run('• '), 9.3, color=NAVY)
    hyperlink(p, name, url, True, 9.3)
    set_font(p.add_run(' — ' + stack + '. '), 9.1, color=GRAY, italic=True)
    set_font(p.add_run(description), 9.3)
    return p


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin = sec.bottom_margin = Inches(0.48)
sec.left_margin = sec.right_margin = Inches(0.58)
sec.header_distance = sec.footer_distance = Inches(0.2)
normal = doc.styles['Normal']
normal.font.name = 'Arial'
normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
normal.font.size = Pt(9.4)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1.0

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
set_font(p.add_run('EDWARD ESTEBAN SÁNCHEZ MOTTA'), 18, True, NAVY)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
p.paragraph_format.line_spacing = 1.0
set_font(p.add_run('DESARROLLADOR BACKEND  |  ALGORITMOS Y VISUALIZACIÓN DE DATOS'), 10.3, True, GRAY)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0.5)
p.paragraph_format.line_spacing = 1.0
set_font(p.add_run('Armenia, Quindío, Colombia  ·  +57 321 753 9293  ·  '), 9.1, color=GRAY)
hyperlink(p, 'edwar5020@gmail.com', 'mailto:edwar5020@gmail.com', size=9.1)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
p.paragraph_format.line_spacing = 1.0
hyperlink(p, 'github.com/EddSanchezz', 'https://github.com/EddSanchezz', size=9.1)
set_font(p.add_run('  ·  '), 9.1, color=GRAY)
hyperlink(p, 'linkedin.com/in/eddsanchezz', 'https://linkedin.com/in/eddsanchezz', size=9.1)

section(doc, 'Perfil profesional')
body(doc, 'Estudiante de Ingeniería de Sistemas y Computación y desarrollador de software con proyectos académicos y personales verificables desde abril de 2024. Enfocado en backend con Java (Spring Boot) y Python (FastAPI, Flask), APIs REST, lógica de negocio y procesamiento de datos. Conocimientos sólidos en algoritmos y análisis de complejidad, y fortalezas en visualización de datos con Power BI y Tableau. Complementa su perfil con frontend (Angular, React), desarrollo móvil Android (Kotlin, Jetpack Compose) y prácticas de despliegue/CI-CD.', 9.4, 1.2)

section(doc, 'Competencias técnicas')
skill(doc, 'Lenguajes', 'Java, Python, Kotlin, TypeScript, JavaScript, SQL, HTML y CSS')
skill(doc, 'Backend y datos', 'Spring Boot, FastAPI, Flask, Node.js/Express, APIs REST, SQLAlchemy, JWT; PostgreSQL, MySQL y SQLite')
skill(doc, 'Algoritmos', 'Ordenamiento y búsqueda (KMP, Boyer–Moore, Rabin–Karp); similitud (Levenshtein, TF-IDF, Jaccard, coseno, Pearson, DTW); clustering jerárquico y embeddings')
skill(doc, 'Visualización y análisis', 'Power BI, Tableau (uso práctico), Pandas, ETL, Streamlit y Chart.js')
skill(doc, 'Frontend y móvil', 'Angular, React, PrimeNG; Android con Jetpack Compose y Hilt')
skill(doc, 'DevOps y herramientas', 'Docker, Docker Compose, GitHub Actions, Dokploy, Multipass, Linux/Ubuntu, Git, Postman y Scrum')

section(doc, 'Experiencia en desarrollo de software')
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
p.paragraph_format.line_spacing = 1.0
keep(p, True)
set_font(p.add_run('Proyectos académicos, personales y técnicos'), 9.6, True)
set_font(p.add_run('  |  abr. 2024 – actualidad'), 9.3, color=GRAY)

project(doc, 'Investment Algorithm', 'https://github.com/EddSanchezz/investment-algorithm', 'Python, Flask, Docker', 'Pipeline ETL para más de 5 años de precios de 23 activos; implementa 12 algoritmos de ordenamiento, similitud (Pearson, DTW y coseno), API REST de 8 endpoints, reportes PDF y 47 pruebas unitarias.')
project(doc, 'Comparador de abstracts científicos', 'https://github.com/EddSanchezz/comparador-textos-algoritmico', 'Python, Streamlit, NLP', 'Compara textos con seis medidas, embeddings y clustering jerárquico; incluye búsqueda de cadenas con KMP, Boyer–Moore y Rabin–Karp.')
project(doc, 'FlyTrack — AeroPuerto Smart', 'https://github.com/EddSanchezz/Practica-DevOps-AeroPuerto-Smart', 'FastAPI, SQLAlchemy, PostgreSQL, Angular, Docker', 'Sistema de vuelos y equipaje con API REST y autenticación JWT; integra Docker Compose y flujos de CI/CD para lint, pruebas y construcción de imágenes.')
project(doc, 'Ticketopia', 'https://github.com/EddSanchezz/ProyectoFinalProgramacion3', 'Java, Maven', 'Sistema de boletería con autenticación, gestión de usuarios, eventos y locaciones, transacciones y persistencia XML; primer proyecto de alcance amplio del portafolio (abr. 2024).')
project(doc, 'ServiLocal', 'https://github.com/EddSanchezz/servi-local', 'Kotlin, Android, Jetpack Compose, Hilt', 'Aplicación móvil con arquitectura por capas, casos de uso, ViewModels/StateFlow y pruebas unitarias; prototipo funcional con datos de demostración.')

section(doc, 'Experiencia laboral e investigación')
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0.5)
p.paragraph_format.line_spacing = 1.0
keep(p, True)
set_font(p.add_run('Apoyo administrativo temporal (temporada de agosto)  ·  ICFES'), 9.5, True)
set_font(p.add_run('  |  ago. 2025 y ago. 2026'), 9.2, color=GRAY)
body(doc, 'Gestionó registros en plataformas institucionales, apoyó la logística de procesos de evaluación y atendió usuarios durante las jornadas.', 9.2, 1)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0.5)
p.paragraph_format.line_spacing = 1.0
keep(p, True)
set_font(p.add_run('Semillero de Investigación en Web Scraping  ·  Celuweb'), 9.5, True)
set_font(p.add_run('  |  jun. – dic. 2024'), 9.2, color=GRAY)
body(doc, 'Desarrolló y optimizó scrapers en Python con Selenium, Scrapy y Beautiful Soup; organizó resultados en JSON, CSV y bases de datos, respetando robots.txt y términos de servicio.', 9.2, 1)

section(doc, 'Educación, certificación e idiomas')
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0.5)
p.paragraph_format.line_spacing = 1.0
set_font(p.add_run('Ingeniería de Sistemas y Computación  ·  Universidad del Quindío'), 9.5, True)
set_font(p.add_run('  |  2022 – actualidad'), 9.2, color=GRAY)
body(doc, 'Cursos relevantes: Análisis de Algoritmos, Estructuras de Datos, Lenguajes Formales e Infraestructura de TI.  ·  Certificación en Spring Boot (Udemy, ago. 2024).  ·  Español nativo; inglés B1.', 9.15, 0)

doc.core_properties.title = 'Hoja de vida — Edward Esteban Sánchez Motta'
doc.core_properties.subject = 'Desarrollador backend, algoritmos y visualización de datos'
doc.core_properties.author = 'Edward Esteban Sánchez Motta'
doc.core_properties.keywords = 'Backend, Java, Python, algoritmos, Power BI, Tableau, FastAPI, Spring Boot'
TEMP_OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(TEMP_OUT)

# Validate the generated file before replacing the requested document.
check = Document(TEMP_OUT)
all_text = '\n'.join(p.text for p in check.paragraphs)
required = ['abril de 2024', 'Power BI', 'Tableau', 'Multipass', 'Dokploy', '47 pruebas unitarias', 'ago. 2025 y ago. 2026']
missing = [x for x in required if x not in all_text]
if missing:
    raise SystemExit('Validación fallida; faltan: ' + ', '.join(missing))
if not TARGET.exists():
    raise SystemExit('No se encontró el archivo original: ' + str(TARGET))
if not BACKUP.exists():
    shutil.copy2(TARGET, BACKUP)
else:
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    BACKUP = TARGET.with_name(TARGET.stem + '_original_backup_' + stamp + TARGET.suffix)
    shutil.copy2(TARGET, BACKUP)
os.replace(TEMP_OUT, TARGET)
print('UPDATED:', TARGET)
print('BACKUP:', BACKUP)
print('PARAGRAPHS:', len(check.paragraphs))
print('DOCX_BYTES:', TARGET.stat().st_size)
