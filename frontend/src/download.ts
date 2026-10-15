export function downloadFile(file: { blob: Blob; filename: string }) {
  const url = URL.createObjectURL(file.blob)
  const link = document.createElement('a')
  link.href = url; link.download = file.filename
  document.body.append(link); link.click(); link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
