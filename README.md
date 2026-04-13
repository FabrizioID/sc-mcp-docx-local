# SC MCP DOCX Local

MCP server local para leer y reemplazar texto en archivos `.docx` usando Microsoft Word via COM automation en Windows.

## Requisitos

- Windows
- Microsoft Word instalado
- Node.js 18+

## Instalacion

```powershell
npm install
```

## Uso

Servidor MCP por `stdio`:

```powershell
node .\server.js
```

## Tools

- `read_docx`
- `replace_text_docx`

## Notas

- Solo funciona con archivos `.docx`
- Crea copias editadas terminadas en `.edited.docx` cuando no se especifica `output_path`
- Depende de Word local, no de Microsoft 365 en la nube
