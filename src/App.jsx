import { useMemo, useState } from "react";
import "./App.css";

const correctionOptions = [
  { id: "integral", title: "Corrección integral", detail: "Recomendada" },
  { id: "verbos", title: "Tiempos verbales", detail: "Presente" },
  { id: "ortografia", title: "Ortografía y puntuación", detail: "" },
  { id: "redaccion", title: "Redacción y coherencia", detail: "" },
  { id: "repeticiones", title: "Evitar repeticiones", detail: "" },
  { id: "pedagogico", title: "Lenguaje pedagógico", detail: "" },
];

const institutionalCriteria = [
  "Redacción preferentemente en tiempo presente.",
  "Lenguaje claro, respetuoso y positivo.",
  "Sin etiquetas ni expresiones descalificadoras.",
  "Sin apreciaciones personales ni valoraciones subjetivas.",
  "Utiliza descripciones objetivas y basadas en hechos observables.",
  "Respeta las ideas originales de la docente.",
  "No inventa aprendizajes ni características.",
];

export default function App() {
  const [original, setOriginal] = useState("");
  const [corrected, setCorrected] = useState("");
  const [selected, setSelected] = useState(["integral"]);
  const [notice, setNotice] = useState("");
  const [loading, setLoading] = useState(false);
  const [modifications, setModifications] = useState([]);

const [archivoWord, setArchivoWord] = useState(null);
const [resultadosWord, setResultadosWord] = useState([]);
const [grado, setGrado] = useState("4.º grado");
const [procesandoWord, setProcesandoWord] = useState(false);

  const count = useMemo(() => original.length, [original]);

  function toggleOption(id) {
    setSelected((current) => {
      if (id === "integral") {
        return current.includes(id) ? [] : ["integral"];
      }

      const withoutIntegral = current.filter(
        (item) => item !== "integral"
      );

      return withoutIntegral.includes(id)
        ? withoutIntegral.filter((item) => item !== id)
        : [...withoutIntegral, id];
    });
  }

  function clearAll() {
    setOriginal("");
    setCorrected("");
    setNotice("");
    setModifications([]);
    setSelected(["integral"]);
  }

  async function copyResult() {
    if (!corrected) {
      setNotice("Todavía no hay una semblanza corregida para copiar.");
      return;
    }

    try {
      await navigator.clipboard.writeText(corrected);
      setNotice("Texto copiado al portapapeles.");
    } catch {
      setNotice(
        "No se pudo copiar automáticamente. Seleccioná el texto y copialo."
      );
    }
  }


async function downloadResult() {
  if (!corrected) {
    setNotice("Primero necesitás procesar una semblanza.");
    return;
  }

  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel
  } = await import("docx");

  const children = [];

  if (resultadosWord.length > 0) {
    resultadosWord.forEach((item) => {
      // Nombre del estudiante destacado
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [
            new TextRun({
              text: item.estudiante,
              bold: true,
              color: "4472C4",
              size: 30
            })
          ],
          spacing: { before: 240, after: 120 }
        })
      );

      // Semblanza corregida
      item.textoCorregido.split("\n").forEach((linea) => {
        if (linea.trim()) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({ text: linea })
              ],
              spacing: { after: 100 }
            })
          );
        }
      });
    });
  } else {
    corrected.split("\n").forEach((linea) => {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: linea })],
          spacing: { after: 160 }
        })
      );
    });
  }

  const documento = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(documento);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "semblanzas-corregidas.docx";
  link.click();

  URL.revokeObjectURL(url);
}


  async function startCorrection() {
    if (!original.trim()) {
      setNotice("Pegá o escribí una semblanza para comenzar.");
      return;
    }

    if (selected.length === 0) {
      setNotice("Seleccioná al menos un criterio de corrección.");
      return;
    }

    setLoading(true);
    setNotice("Procesando la semblanza...");
    setModifications([]);
    setCorrected("");

    try {
      const response = await fetch(
        "http://localhost:8080/api/semblanzas/corregir",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            texto: original,
            criterios: selected.map((item) => item.toUpperCase()),
          }),
        }
      );

      if (!response.ok) {
        throw new Error("El servidor devolvió un error.");
      }

      const data = await response.json();

      setCorrected(data.textoCorregido || "");
      setModifications(data.modificaciones || []);
      setNotice("La respuesta del servidor se recibió correctamente.");
    } catch (error) {
      console.error("Error al corregir:", error);

      setNotice(
        "No se pudo conectar con el servidor. Verificá que el backend esté iniciado."
      );
    } finally {
      setLoading(false);
    }
  }

  async function procesarWord() {
  if (!archivoWord) {
    setNotice("Primero seleccioná un documento Word.");
    return;
  }

  if (!archivoWord.name.toLowerCase().endsWith(".docx")) {
    setNotice("El archivo debe tener formato .docx.");
    return;
  }

  setProcesandoWord(true);
  setResultadosWord([]);
  setNotice("Procesando las semblanzas del documento...");

  try {
    const formulario = new FormData();
    formulario.append("archivo", archivoWord);

    const response = await fetch(
      "http://localhost:8080/api/semblanzas/corregir-word",
      {
        method: "POST",
        body: formulario,
      }
    );

    if (!response.ok) {
      const detalle = await response.text();
      throw new Error(detalle || "No se pudo procesar el documento.");
    }

    const data = await response.json();

    if (!Array.isArray(data) || data.length === 0) {
      throw new Error("No se encontraron semblanzas en el documento.");
    }

    setResultadosWord(data);

    setOriginal(
      data.map(item =>
        `${item.estudiante}\n${item.textoOriginal}`
      ).join("\n\n")
    );

    setCorrected(
      data.map(item =>
        `${item.estudiante}\n${item.textoCorregido}`
      ).join("\n\n")
    );

    setModifications(
      data.flatMap(item => item.modificaciones || [])
    );

    setNotice(
      `Se procesaron ${data.length} semblanzas del documento.`
    );

  } catch (error) {
    console.error("Error al procesar Word:", error);
    setNotice(`No se pudo procesar el Word: ${error.message}`);

  } finally {
    setProcesandoWord(false);
  }
}

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <img
            className="school-crest"
            src="/escudo-escuela.png"
            alt="Escudo de la Escuela General José de San Martín"
          />

          <div className="brand-divider" />

          <div className="brand-copy">
            <div className="brand-title-row">
              <span className="brand-icon">▤</span>

              <div>
                <h1>
                  Corrector de
                  <br className="title-break" /> Semblanzas Escolares
                </h1>

                <p>
                  Herramienta de apoyo para la redacción de informes y
                  semblanzas
                </p>
              </div>
            </div>

            <span className="school-name">
              ESCUELA GENERAL JOSÉ DE SAN MARTÍN · TANCACHA
            </span>
          </div>

          <div className="header-decoration" aria-hidden="true">
            <span>✎</span>
            <span>✦</span>
            <span>▱</span>
          </div>
        </div>
      </header>

      <div className="workspace">
        <aside className="sidebar">
          <div className="side-mark">✿</div>

          <nav aria-label="Navegación principal">
            <button className="nav-item active">
              <span>⌂</span>Inicio
            </button>

            <button
              className="nav-item"
              onClick={() =>
                setNotice(
                  "La configuración institucional estará disponible en una próxima etapa."
                )
              }
            >
              <span>⚙</span>Configuración
            </button>

            <button
              className="nav-item"
              onClick={() =>
                setNotice(
                  "Los criterios institucionales están visibles en esta pantalla."
                )
              }
            >
              <span>ⓘ</span>Criterios
            </button>

            <button
              className="nav-item"
              onClick={() =>
                setNotice(
                  "Ayuda: escribí o pegá el texto, elegí qué revisar y luego solicitá la corrección."
                )
              }
            >
              <span>?</span>Ayuda
            </button>
          </nav>

          <div className="sidebar-doodle" aria-hidden="true">
            ♡
          </div>
        </aside>

        <main className="main-content">
          <section className="welcome-card">
            <div className="welcome-art" aria-hidden="true">
              <span className="book book-back" />
              <span className="book book-front" />
              <span className="spark">✦</span>
            </div>

            <div>
              <h2>¡Bienvenida!</h2>

              <p>
                Cargá la semblanza, elegí el tipo de corrección y obtené una
                versión mejorada, respetando el sentido de tu texto y los
                criterios institucionales.
              </p>
            </div>
          </section>

          {/* TRES VENTANAS DE REVISIÓN */}

          <section
            className="editors-grid three-editors"
            aria-label="Revisión de semblanzas"
          >
            {/* VENTANA 1: ORIGINAL */}

            <article className="editor-card original-card">
              <div className="card-heading">
                <span className="heading-icon document-icon">▤</span>

                <div>
                  <h3>1. Semblanza original</h3>
                  <p>Texto enviado para revisar.</p>
                </div>
              </div>

<div className="word-upload">
  <label htmlFor="archivo-word">
    <span>📄 Cargar semblanzas desde Word (.docx)</span>
  </label>

  <input
    id="archivo-word"
    type="file"
    accept=".docx"
    onChange={(event) =>
      setArchivoWord(event.target.files?.[0] || null)
    }
  />

  {archivoWord && (
    <p>Archivo seleccionado: {archivoWord.name}</p>
  )}

  <button
    type="button"
    onClick={procesarWord}
    disabled={!archivoWord || procesandoWord}
  >
    {procesandoWord ? "Procesando..." : "Procesar documento Word"}
  </button>
</div>

           <button
  type="button"
  className="secondary-button"
  onClick={() => window.location.reload()}
  style={{ marginBottom: "12px" }}
>
  ↻ Nueva revisión
</button>

<textarea
  value={original}
  onChange={(event) => {
    setOriginal(event.target.value);
    setNotice("");
  }}
  placeholder="Escribí o pegá la semblanza del estudiante..."
  maxLength={5000}
  aria-label="Semblanza original"
/>  

                

              <div className="editor-footer">
                <span>Texto original</span>
                <strong>{count} / 5000</strong>
              </div>
            </article>

            {/* VENTANA 2: CORREGIDA */}

            <article className="editor-card corrected-card">
              <div className="card-heading">
                <span className="heading-icon check-icon">✓</span>

                <div>
                  <h3>2. Semblanza corregida</h3>
                  <p>Resultado de la revisión.</p>
                </div>
              </div>

              <textarea
                value={corrected}
                onChange={(event) => setCorrected(event.target.value)}
                placeholder="Aquí aparecerá la versión corregida..."
                aria-label="Semblanza corregida"
              />

              <div className="editor-footer">
                <span>Podés revisar y editar el resultado.</span>
                <strong>{corrected.length} caracteres</strong>
              </div>
            </article>

            {/* VENTANA 3: MODIFICACIONES */}

            <article className="editor-card modifications-card">
              <div className="card-heading">
                <span className="heading-icon">☷</span>

                <div>
                  <h3>3. ¿Qué se corrigió?</h3>
                  <p>Detalle de los cambios realizados.</p>
                </div>
              </div>

              {modifications.length > 0 ? (
                <ul className="modifications-list">
                  {modifications.map((item, index) => (
                    <li key={index}>
                      <span>✓</span>
                      {item}
                    </li>
                  ))}
                </ul>
              ) : (
                <p className="empty-modifications">
                  Aquí aparecerá el detalle de las correcciones realizadas.
                </p>
              )}
            </article>
          </section>

          <section className="lower-grid">
            <div className="options-section">
              <h3 className="section-title">
                <span>✧</span> ¿Qué querés corregir?
              </h3>

              <div className="options-grid">
                {correctionOptions.map((option) => (
                  <label
                    className={`option-tile ${
                      selected.includes(option.id) ? "chosen" : ""
                    }`}
                    key={option.id}
                  >
                    <input
                      type="checkbox"
                      checked={selected.includes(option.id)}
                      onChange={() => toggleOption(option.id)}
                    />

                    <span className="custom-check">
                      {selected.includes(option.id) ? "✓" : ""}
                    </span>

                    <span className="option-text">
                      {option.title}
                      {option.detail && <small>({option.detail})</small>}
                    </span>
                  </label>
                ))}
              </div>
            </div>

            <aside className="criteria-card">
              <div className="criteria-heading">
                <span>◎</span>
                <h3>Criterios institucionales</h3>
              </div>

              <ul>
                {institutionalCriteria.map((item) => (
                  <li key={item}>
                    <span>✓</span>
                    {item}
                  </li>
                ))}
              </ul>

              <div className="pencil-cup" aria-hidden="true">
                <span>✎</span>
                <span>♥</span>
              </div>
            </aside>
          </section>

          <div className="action-area">
            <div className="actions">
              <button
                className="primary-button"
                onClick={startCorrection}
                disabled={loading}
              >
                <span>✧</span>
                {loading ? "Procesando..." : "Corregir semblanza"}
              </button>

              <button className="secondary-button" onClick={copyResult}>
                <span>▢</span> Copiar resultado
              </button>

              <button className="secondary-button" onClick={downloadResult}>
                <span>⇩</span> Descargar
              </button>

              <button
                className="secondary-button clear-button"
                onClick={clearAll}
              >
                <span>♲</span> Limpiar
              </button>
            </div>

            {notice && (
              <p className="notice" role="status">
                {notice}
              </p>
            )}
          </div>

          <footer className="page-footer">
            Una herramienta de acompañamiento a la escritura docente · Escuela
            General José de San Martín
          </footer>
        </main>
      </div>
    </div>
  );
}