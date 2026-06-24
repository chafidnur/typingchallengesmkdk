/**
 * Fungsi DoGet (Entry point Utama Web App GAS)
 */
function doGet(e) {
  // PERBAIKAN KRITIS: Selalu muat "Index" sebagai wadah utama layout SPA.
  // Jangan langsung memuat halaman lain agar CSS dan JS global tetap terbawa.
  let html = HtmlService.createTemplateFromFile("resources/views/layouts/Index");
  return html.evaluate()
             .setTitle("Typing Challenge SMK Duta Karya")
             .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
             .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi untuk Include file komponen terpisah (CSS, JS, Utils) ke dalam Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fungsi untuk menyuplai HTML dinamis saat pindah halaman secara SPA (Client AJAX)
 */
function getHtmlTemplate(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}