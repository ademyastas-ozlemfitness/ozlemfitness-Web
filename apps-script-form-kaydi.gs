/**
 * Özlem Fitness — form başvurusu kaydı
 *
 * Bu kodu Google Apps Script düzenleyicisine yapıştır (mevcut kodun yerine),
 * kaydet, sonra Dağıt → Dağıtımı yönet → düzenle → yeni sürüm → dağıt.
 * Dağıtım adresi (/exec) DEĞİŞMEZ, yani sitedeki adrese dokunmaya gerek yok.
 *
 * Ne değişti:
 * - Başvuru kimliği (basvuru_id) kaydediliyor → CRM mükerrer kaydı ayırt eder
 * - UTM alanları, yönlendiren sayfa ve ilk sayfa kaydediliyor → lead'in
 *   hangi reklamdan/kaynaktan geldiği belli olur
 * - Meta tıklama kimlikleri (fbclid, _fbp, _fbc) kaydediliyor → ileride
 *   Conversions API ile "bu kişi üye oldu" bilgisi Meta'ya geri gönderilebilir
 * - IP adresi kaydediliyor → İYS web onayında zorunlu alan
 *
 * Başlık satırı tabloda yoksa ilk çalıştırmada kendiliğinden yazılır.
 * Mevcut tabloda eski başlıklar duruyorsa: yeni sütunları sağa eklemek için
 * tabloyu boşalt ya da başlık satırını elle tamamla (sıra aşağıdaki gibi).
 */

var SPREADSHEET_ID = '1-uoC7MkGRYOOvMLcqaPyJiXREalBm0MtKDA2QKdvfv0';
var SAYFA_ADI = 'Form Başvuruları';

var BASLIKLAR = [
  'Tarih-Saat',
  'Başvuru ID',
  'Ad',
  'Telefon',
  'Doğum Günü',
  'E-posta',
  'Hangi Butondan Geldi',
  'Aydınlatma Okundu',
  'Ticari İleti İzni',
  'IP Adresi',
  'utm_source',
  'utm_medium',
  'utm_campaign',
  'utm_content',
  'utm_term',
  'Yönlendiren',
  'İlk Sayfa',
  'fbclid',
  'fbp',
  'fbc'
];

var AYLAR = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
             'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];

function doPost(e) {
  try {
    var veri = JSON.parse(e.postData.contents);
    var sayfa = sayfayiAl();

    sayfa.appendRow([
      veri.gonderim_tarihi ? new Date(veri.gonderim_tarihi) : new Date(),
      veri.basvuru_id || '',
      veri.ad || '',
      "'" + (veri.telefon || ''),          // başta sıfır kaybolmasın diye metin
      dogumGunu(veri.dogum_gun, veri.dogum_ay),
      veri.eposta || '',
      veri.kaynak || '',
      veri.aydinlatma_okundu || '',
      veri.ticari_ileti_izni || '',
      veri.ip_adresi || '',
      veri.utm_source || '',
      veri.utm_medium || '',
      veri.utm_campaign || '',
      veri.utm_content || '',
      veri.utm_term || '',
      veri.yonlendiren || '',
      veri.ilk_sayfa || '',
      veri.fbclid || '',
      veri.fbp || '',
      veri.fbc || ''
    ]);

    if (veri.eposta) {
      try { epostaOnayGonder(veri.eposta, veri.ad); } catch (postaHata) { /* mail gitmese de kayıt zaten yapıldı */ }
    }

    return cevap({ ok: true });
  } catch (hata) {
    return cevap({ ok: false, hata: String(hata) });
  }
}

function epostaOnayGonder(eposta, ad) {
  var isim = ad ? ad.split(' ')[0] : '';
  var konu = 'Başvurun Alındı — Özlem Fitness';
  var govde =
    (isim ? 'Merhaba ' + isim + ',' : 'Merhaba,') + '\n\n' +
    'Özlem Fitness üzerinden gönderdiğin ücretsiz analiz başvurusu bize ulaştı.\n' +
    'Cengiz Hoca en kısa sürede seni arayacak.\n\n' +
    'Bu süreçte bir sorun yaşarsan 0536 584 80 18 numarasından bize ulaşabilirsin.\n\n' +
    'Özlem Fitness Center\n' +
    'İzzettin Çalışlar Cad. No: 31/B, Bahçelievler / İstanbul';
  MailApp.sendEmail({ to: eposta, subject: konu, body: govde });
}

function sayfayiAl() {
  var kitap = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sayfa = kitap.getSheetByName(SAYFA_ADI) || kitap.getSheets()[0];

  // Başlık satırı yoksa yaz
  if (sayfa.getLastRow() === 0) {
    sayfa.appendRow(BASLIKLAR);
    sayfa.getRange(1, 1, 1, BASLIKLAR.length).setFontWeight('bold');
    sayfa.setFrozenRows(1);
  }
  return sayfa;
}

function dogumGunu(gun, ay) {
  if (!gun || !ay) return '';
  var ayNo = parseInt(ay, 10);
  var ayAdi = (ayNo >= 1 && ayNo <= 12) ? AYLAR[ayNo - 1] : ay;
  return gun + ' ' + ayAdi;
}

function cevap(nesne) {
  return ContentService
    .createTextOutput(JSON.stringify(nesne))
    .setMimeType(ContentService.MimeType.JSON);
}
