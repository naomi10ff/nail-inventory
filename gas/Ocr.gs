/**
 * 商品ラベルの写真からテキストを読み取る(Google DriveのOCR変換機能を利用)。
 * このファイルを使うには、Apps Scriptエディタで「サービス」から
 * 「Drive API」(Advanced Google Services)を追加で有効にする必要がある。
 * 詳細はREADME参照。
 */

function ocrProductLabel_(session, p) {
  if (!p.imageBase64) throw new Error('画像データがありません');

  var mimeType = p.mimeType || 'image/jpeg';
  var bytes = Utilities.base64Decode(p.imageBase64);
  var blob = Utilities.newBlob(bytes, mimeType, 'label.jpg');

  var fileMetadata = {
    title: 'ocr-temp-' + Utilities.getUuid(),
    mimeType: mimeType
  };

  var file = Drive.Files.insert(fileMetadata, blob, { ocr: true, ocrLanguage: 'ja' });
  var text = '';
  try {
    var doc = DocumentApp.openById(file.id);
    text = doc.getBody().getText();
  } finally {
    Drive.Files.remove(file.id);
  }

  var lines = text
    .split('\n')
    .map(function (line) { return line.trim(); })
    .filter(function (line) { return line.length > 0; });

  return {
    rawText: text,
    lines: lines,
    guessedName: lines[0] || '',
    guessedBrand: lines[1] || ''
  };
}
