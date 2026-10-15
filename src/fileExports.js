function browserDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

export async function chooseExportDestination({ browser = window, download = browserDownload } = {}) {
  if (typeof browser.showDirectoryPicker !== "function") {
    const proceed = browser.confirm('This browser cannot open an export folder picker. To choose a destination for every download, enable "Ask where to save each file before downloading" in your browser settings.\n\nContinue with a browser download?');
    return proceed ? { save: async (blob, filename) => { download(blob, filename); return true; } } : null;
  }

  let directory;
  try {
    // Request access on the original click, before asynchronous report generation consumes user activation.
    directory = await browser.showDirectoryPicker({ id: "exam-exports", mode: "readwrite" });
  } catch (error) {
    if (error.name === "AbortError") return null;
    throw error;
  }

  return { save: async (blob, filename) => {
    let file;
    try {
      file = await directory.getFileHandle(filename);
    } catch (error) {
      if (error.name !== "NotFoundError") throw error;
    }
    if (file && !browser.confirm(`"${filename}" already exists in "${directory.name}". Replace it?`)) return false;
    file ??= await directory.getFileHandle(filename, { create: true });
    const writable = await file.createWritable();
    try {
      await writable.write(blob);
      await writable.close();
    } catch (error) {
      try { await writable.abort(); } catch { /* Preserve the original write error. */ }
      throw error;
    }
    return true;
  } };
}
