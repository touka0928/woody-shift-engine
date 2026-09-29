/**
 * ウッディーのウキウキシフト
 * 自動更新ブートストラップ v1.0.0
 *
 * 初回のみ、このファイルを Code.gs に貼り付けます。
 * 以後のエンジン更新は GitHub から自動取得します。
 */

const WOODY_BOOT = Object.freeze({
  VERSION: 'v1.0.0',
  MANIFEST_URL: 'https://raw.githubusercontent.com/touka0928/woody-shift-engine/main/manifest.json',
  ALLOWED_ENGINE_URL: 'https://raw.githubusercontent.com/touka0928/woody-shift-engine/main/engine.gs',
  CACHE_CHUNK_SIZE: 7000
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('シフト自動作成')
    .addItem('✨ まとめてシフト作成', 'woodyShiftRunOneClick')
    .addSeparator()
    .addItem('① 初期設定を確認', 'woodyShiftValidateInitialSetup')
    .addItem('高度ルールを確認', 'woodyShiftValidateAdvancedRules')
    .addSeparator()
    .addItem('② 月間シフトを作成', 'woodyShiftGenerate')
    .addItem('③ シフトを再最適化', 'woodyShiftReoptimize')
    .addItem('④ 現在シフトを検証', 'woodyShiftValidateCurrent')
    .addItem('⑤ 最終確定チェック', 'woodyShiftFinalCheck')
    .addSeparator()
    .addItem('シフトを確定', 'woodyShiftConfirmCurrent')
    .addItem('確定後の振替を反映', 'woodyShiftApplyConfirmedChanges')
    .addItem('次月ファイルを作成', 'woodyShiftCreateNextMonthFile')
    .addSeparator()
    .addItem('指定日のみ再生成', 'woodyShiftRegenerateDay')
    .addItem('指定職員のみ再生成', 'woodyShiftRegenerateStaff')
    .addItem('指定単位のみ再生成', 'woodyShiftRegenerateUnit')
    .addItem('当日欠勤を反映', 'woodyShiftRegisterAbsence')
    .addSeparator()
    .addItem('試算用コピーを作成', 'woodyShiftCreateTrialCopy')
    .addItem('⑥ 直前状態へ戻す', 'woodyShiftRestoreBackup')
    .addSeparator()
    .addItem('🔄 エンジン情報', 'woodyEngineInfo')
    .addToUi();
}

function woodyShiftRunOneClick() { woodyRun_('shiftRunOneClick'); }
function woodyShiftValidateInitialSetup() { woodyRun_('shiftValidateInitialSetup'); }
function woodyShiftValidateAdvancedRules() { woodyRun_('shiftValidateAdvancedRules'); }
function woodyShiftGenerate() { woodyRun_('shiftGenerate'); }
function woodyShiftReoptimize() { woodyRun_('shiftReoptimize'); }
function woodyShiftValidateCurrent() { woodyRun_('shiftValidateCurrent'); }
function woodyShiftFinalCheck() { woodyRun_('shiftFinalCheck'); }
function woodyShiftConfirmCurrent() { woodyRun_('shiftConfirmCurrent'); }
function woodyShiftApplyConfirmedChanges() { woodyRun_('shiftApplyConfirmedChanges'); }
function woodyShiftCreateNextMonthFile() { woodyRun_('shiftCreateNextMonthFile'); }
function woodyShiftRegenerateDay() { woodyRun_('shiftRegenerateDay'); }
function woodyShiftRegenerateStaff() { woodyRun_('shiftRegenerateStaff'); }
function woodyShiftRegenerateUnit() { woodyRun_('shiftRegenerateUnit'); }
function woodyShiftRegisterAbsence() { woodyRun_('shiftRegisterAbsence'); }
function woodyShiftCreateTrialCopy() { woodyRun_('shiftCreateTrialCopy'); }
function woodyShiftRestoreBackup() { woodyRun_('shiftRestoreBackup'); }

function woodyRun_(entryPoint) {
  const allowed = [
    'shiftRunOneClick','shiftValidateInitialSetup','shiftValidateAdvancedRules',
    'shiftGenerate','shiftReoptimize','shiftValidateCurrent','shiftFinalCheck',
    'shiftConfirmCurrent','shiftApplyConfirmedChanges','shiftCreateNextMonthFile',
    'shiftRegenerateDay','shiftRegenerateStaff','shiftRegenerateUnit',
    'shiftRegisterAbsence','shiftCreateTrialCopy','shiftRestoreBackup'
  ];
  if (allowed.indexOf(entryPoint) < 0) throw new Error('許可されていない処理です。');
  const loaded = woodyLoadEngine_();
  eval(loaded.code + '\n;' + entryPoint + '();');
}

function woodyLoadEngine_() {
  try {
    const manifestResponse = UrlFetchApp.fetch(WOODY_BOOT.MANIFEST_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (manifestResponse.getResponseCode() !== 200) {
      throw new Error('更新情報の取得に失敗しました。HTTP ' + manifestResponse.getResponseCode());
    }
    const manifest = JSON.parse(manifestResponse.getContentText('UTF-8'));
    if (!manifest || manifest.engine_url !== WOODY_BOOT.ALLOWED_ENGINE_URL) {
      throw new Error('更新元URLが正しくありません。');
    }
    const engineResponse = UrlFetchApp.fetch(manifest.engine_url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (engineResponse.getResponseCode() !== 200) {
      throw new Error('エンジン取得に失敗しました。HTTP ' + engineResponse.getResponseCode());
    }
    const code = engineResponse.getContentText('UTF-8');
    if (code.length < 50000 || code.indexOf("const SHIFT_APP") < 0) {
      throw new Error('取得したエンジンの内容が不正です。');
    }
    const versionMatch = code.match(/VERSION:\s*'([^']+)'/);
    const codeVersion = versionMatch ? versionMatch[1] : '';
    if (!codeVersion || codeVersion !== String(manifest.version || '')) {
      throw new Error('エンジンのバージョン情報が一致しません。');
    }
    if (manifest.sha256 && woodySha256_(code) !== String(manifest.sha256).toLowerCase()) {
      throw new Error('エンジンの整合性チェックに失敗しました。');
    }
    woodyCacheEngine_(code, codeVersion);
    return {code: code, version: codeVersion, source: '最新版'};
  } catch (error) {
    const cached = woodyReadCachedEngine_();
    if (cached) return {code: cached.code, version: cached.version, source: '前回正常版'};
    throw new Error('ウッディーのエンジンを読み込めませんでした。\n' + (error && error.message ? error.message : error));
  }
}

function woodyEngineInfo() {
  try {
    const loaded = woodyLoadEngine_();
    SpreadsheetApp.getUi().alert(
      'ウッディー エンジン情報',
      'ブートストラップ：' + WOODY_BOOT.VERSION + '\n' +
      'エンジン：' + loaded.version + '\n' +
      '読込元：' + loaded.source + '\n\n' +
      'エンジンは自動更新されます。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('エンジン情報', String(error.message || error), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function woodySha256_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    const n = b < 0 ? b + 256 : b;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}

function woodyCacheEngine_(code, version) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('WOODY_ENGINE_VERSION') === version &&
      Number(props.getProperty('WOODY_ENGINE_CHUNKS') || 0) > 0) return;
  const oldCount = Number(props.getProperty('WOODY_ENGINE_CHUNKS') || 0);
  for (let i = 0; i < oldCount; i++) props.deleteProperty('WOODY_ENGINE_' + i);
  const chunks = [];
  for (let i = 0; i < code.length; i += WOODY_BOOT.CACHE_CHUNK_SIZE) {
    chunks.push(code.substring(i, i + WOODY_BOOT.CACHE_CHUNK_SIZE));
  }
  const values = {
    WOODY_ENGINE_VERSION: version,
    WOODY_ENGINE_CHUNKS: String(chunks.length)
  };
  chunks.forEach(function(chunk, i) { values['WOODY_ENGINE_' + i] = chunk; });
  props.setProperties(values, false);
}

function woodyReadCachedEngine_() {
  const props = PropertiesService.getDocumentProperties();
  const count = Number(props.getProperty('WOODY_ENGINE_CHUNKS') || 0);
  const version = String(props.getProperty('WOODY_ENGINE_VERSION') || '');
  if (!count || !version) return null;
  let code = '';
  for (let i = 0; i < count; i++) {
    const chunk = props.getProperty('WOODY_ENGINE_' + i);
    if (chunk === null) return null;
    code += chunk;
  }
  if (code.length < 50000 || code.indexOf("const SHIFT_APP") < 0) return null;
  return {code: code, version: version};
}
