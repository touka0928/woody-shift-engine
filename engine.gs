/**
 * 放課後等デイサービス向け シフト自動作成アプリ
 * 配布版 v1.2 / GASエンジン v1.1.2
 *
 * 対象:
 *   ウッディーのウキウキシフト（配布用）
 *
 * 方針:
 * - 入力値は原則として表シートから取得する。
 * - 出力用「シフト」シートを直接書き換えず、
 *   内部シート「自動計算」の勤務0/1と所属単位を確定する。
 * - 既存数式による法定配置・送迎・月間集計・★表示を活かす。
 * - 個人固有の隠し条件、特定人物の優先順位は持たない。
 * - エラー・警告は日本語で表示する。
 *
 * 重要:
 * このスクリプトは配置管理支援ツールです。
 * 加算や請求の最終可否は、最新の公式通知・届出内容・自治体運用を確認してください。
 */

const SHIFT_APP = Object.freeze({
  VERSION: 'v1.1.2',
  MAX_STAFF: 15,
  MAX_DAYS: 31,

  SHEETS: Object.freeze({
    SETUP: '初回セットアップ',
    STAFF: '職員セットアップ',
    WORK_STYLE: '働き方設定',
    WORK_PATTERN: '勤務パターン設定',
    SERVICE_TIME: '提供時間設定',
    PLACEMENT: '配置・加算設定',
    REQUEST: '希望入力',
    USAGE: '利用予定',
    SHIFT: 'シフト',
    AUTO: '自動計算',
    MONTHLY: '月間集計',
    VEHICLE: '送迎車両',
    LOCK: '配置ロック',
    FINAL_CHECK: '最終チェック',
    SUGGESTIONS: '修正候補',
    CHANGE_DIFF: '変更比較',
    COST: '人件費設定',
    BALANCE: '負担バランス',
    REVENUE: '収益・加算',
    CONTROL: '管理コンソール',
    REASON: '配置理由',
    ACTUAL: '実績',
    FACILITY_RULE: '施設ルール',
    PERSON_RULE: '個人ルール',
    PAIR_RULE: 'ペアルール',
    EXCEPTION_RULE: '月間・例外ルール',
    CONFIRMED_HISTORY: '_確定履歴',
    ACTUAL_LOG: '_実績変更ログ',
    BACKUP: '_GAS_BACKUP',
    LOG: '_GAS_LOG'
  }),

  // この2つだけはテンプレート内部の「GAS出力口」として固定する。
  // 表側の入力列はヘッダー名で読むため、列の並び替えに比較的強い。
  RANGE: Object.freeze({
    WORK_PLAN: 'P4:AT18',      // 15名 × 31日：勤務=1 / 非勤務=0
    UNIT_PLAN: 'B86:AF100',    // 15名 × 31日：空白=所属 / 第1 / 第2
    SCORE: 'P22:AT36',         // 15名 × 31日：既存数式の候補日スコア
    AUTO_META: 'A3:O18',       // 必要勤務日数・時間等
    HALF_DAY: 'AW4:AX18'       // 単半対象日（送迎・法定配置から除外される日）
  }),

  LOG_HEADERS: [
    '日時', 'レベル', '処理', '対象', 'メッセージ', 'バージョン'
  ]
});


/* =========================================================
 * 1. メニュー
 * ========================================================= */

/**
 * スプレッドシートを開いた時に管理メニューを作る。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('シフト自動作成')
    .addItem('✨ まとめてシフト作成', 'shiftRunOneClick')
    .addSeparator()
    .addItem('① 初期設定を確認', 'shiftValidateInitialSetup')
    .addItem('高度ルールを確認', 'shiftValidateAdvancedRules')
    .addSeparator()
    .addItem('② 月間シフトを作成', 'shiftGenerate')
    .addItem('③ シフトを再最適化', 'shiftReoptimize')
    .addItem('④ 現在シフトを検証', 'shiftValidateCurrent')
    .addItem('⑤ 最終確定チェック', 'shiftFinalCheck')
    .addSeparator()
    .addItem('シフトを確定', 'shiftConfirmCurrent')
    .addItem('確定後の振替を反映', 'shiftApplyConfirmedChanges')
    .addItem('次月ファイルを作成', 'shiftCreateNextMonthFile')
    .addSeparator()
    .addItem('指定日のみ再生成', 'shiftRegenerateDay')
    .addItem('指定職員のみ再生成', 'shiftRegenerateStaff')
    .addItem('指定単位のみ再生成', 'shiftRegenerateUnit')
    .addItem('当日欠勤を反映', 'shiftRegisterAbsence')
    .addSeparator()
    .addItem('試算用コピーを作成', 'shiftCreateTrialCopy')
    .addItem('⑥ 直前状態へ戻す', 'shiftRestoreBackup')
    .addToUi();
}


/* =========================================================
 * 2. 公開関数（メニューから呼ぶ）
 * ========================================================= */

/**
 * 初期設定→高度ルール確認→生成→検証までをまとめて実行する。
 */
function shiftRunOneClick() {
  runSafely_('まとめてシフト作成', function () {
    const ctx = loadContext_();
    const initial = validateInitialSetup_(ctx);
    const advanced = validateAdvancedRuleDefinitions_(ctx);
    const errors = unique_([].concat(initial.errors || [], advanced.errors || []));
    const warnings = unique_([].concat(initial.warnings || [], advanced.warnings || []));
    if (errors.length) {
      writeOperationStatus_(ctx.ss, '🔴 設定修正が必要', '初期設定・高度ルールを修正', errors.slice(0,3).join(' / '));
      showAlert_('まとめてシフト作成', formatValidationResult_({errors:errors, warnings:warnings}));
      return;
    }
    writeOperationStatus_(ctx.ss, '🟡 作成中', '生成後の検証を実行中', warnings.slice(0,3).join(' / '));
    generateShift_({mode:'generate', scope:{type:'all'}});
  });
}

/**
 * 初期設定の不足を確認する。
 */
function shiftValidateInitialSetup() {
  runSafely_('初期設定確認', function () {
    const ctx = loadContext_();
    const result = validateInitialSetup_(ctx);

    writeValidationLog_('初期設定確認', result);

    if (result.errors.length === 0 && result.warnings.length === 0) {
      showAlert_(
        '初期設定確認',
        '初期設定に重大な不足は見つかりませんでした。\n\n' +
        '次に「利用予定」と「希望入力」を入力して、月間シフトを作成してください。'
      );
      return;
    }

    showAlert_('初期設定確認', formatValidationResult_(result));
  });
}


/**
 * 月間シフトを新規生成する。
 */
function shiftGenerate() {
  runSafely_('月間シフト作成', function () {
    generateShift_({
      mode: 'generate',
      scope: { type: 'all' }
    });
  });
}


/**
 * 現在シフトを基準に全体を再最適化する。
 */
function shiftReoptimize() {
  runSafely_('シフト再最適化', function () {
    generateShift_({
      mode: 'reoptimize',
      scope: { type: 'all' }
    });
  });
}


/**
 * 現在のシフトを数式結果とGAS側ロジックの両方から検証する。
 */
function shiftValidateCurrent() {
  runSafely_('現在シフト検証', function () {
    const ctx = loadContext_();
    const plan = readCurrentPlan_(ctx);
    const result = validateCurrentShift_(ctx, plan);

    writeValidationLog_('現在シフト検証', result);

    const suggestions = buildCorrectionSuggestions_(ctx, plan, result, 10);
    writeCorrectionSuggestions_(ctx, suggestions, result);
    let message = formatValidationResult_(result);

    if (suggestions.length > 0) {
      message += '\n\n【修正候補】\n' +
        suggestions.slice(0, 3).map(function (x, i) {
          return (i + 1) + '. ' + x;
        }).join('\n');
    }

    showAlert_('現在シフト検証', message);
  });
}



/**
 * 確定前のハード条件を一括確認する。
 */
function shiftFinalCheck() {
  runSafely_('最終確定チェック', function () {
    const ctx = loadContext_();
    const plan = readCurrentPlan_(ctx);
    const result = validateCurrentShift_(ctx, plan);
    const suggestions = buildCorrectionSuggestions_(ctx, plan, result, 10);
    writeCorrectionSuggestions_(ctx, suggestions, result);
    SpreadsheetApp.flush();

    const sheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.FINAL_CHECK);
    const verdict = String(sheet.getRange('B4').getDisplayValue() || '').trim();
    let message = verdict || '判定できませんでした。';
    if (result.errors.length > 0) {
      message += '\n\n' + formatValidationResult_(result);
    } else if (suggestions.length > 0) {
      message += '\n\n改善候補は「修正候補」シートに表示しました。';
    }
    showAlert_('最終確定チェック', message);
  });
}


/**
 * 高度ルール4シートの参照先・矛盾を確認する。
 */
function shiftValidateAdvancedRules() {
  runSafely_('高度ルール確認', function () {
    const ctx = loadContext_();
    const result = validateAdvancedRuleDefinitions_(ctx);
    showAlert_('高度ルール確認', formatValidationResult_(result));
  });
}

/**
 * 現在シフトを確定し、印刷用シートを値としてスナップショット保存する。
 */
function shiftConfirmCurrent() {
  runSafely_('シフト確定', function () {
    const ctx = loadContext_();
    const plan = readCurrentPlan_(ctx);
    const result = validateCurrentShift_(ctx, plan);
    if (result.errors.length > 0) {
      throw new Error('確定できません。\n\n' + formatValidationResult_(result));
    }

    const monthKey = Utilities.formatDate(
      ctx.targetMonth,
      ctx.ss.getSpreadsheetTimeZone(),
      'yyyyMM'
    );
    const snapName = '_確定_' + monthKey;
    const old = ctx.ss.getSheetByName(snapName);
    if (old) ctx.ss.deleteSheet(old);

    const source = ctx.ss.getSheetByName('シフト_印刷形式') ||
      getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.SHIFT);
    const snap = source.copyTo(ctx.ss).setName(snapName);
    const range = snap.getDataRange();
    range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    snap.hideSheet();

    const hist = ensureConfirmedHistorySheet_(ctx.ss);
    hist.appendRow([
      new Date(),
      ctx.targetMonth,
      snapName,
      result.warnings.length,
      result.warnings.join(' / '),
      Session.getActiveUser().getEmail() || ''
    ]);
    initializeActualSheet_(ctx, plan);
    writeOperationStatus_(ctx.ss, '🟢 確定', '実績入力・確定後変更が可能', 'スナップショット：' + snapName);
    appendLog_('INFO', 'シフト確定', monthKey, snapName);
    showAlert_('シフト確定', '現在のシフトを確定しました。\nスナップショット：' + snapName + '\n\n「実績」シートも準備しました。');
  });
}

/**
 * 「月間・例外ルール」の「確定後振替」を配置ロックへ変換し、再最適化する。
 * 対象=職員名または職員ID、開始日=休みにする日、終了日=出勤する日、値1=第1/第2（任意）。
 */
function shiftApplyConfirmedChanges() {
  runSafely_('確定後振替', function () {
    const ctx = loadContext_();
    const rules = (ctx.rules && ctx.rules.exceptionRules || []).filter(function (r) {
      return isRuleEnabled_(r) && String(r['ルール種別'] || '').trim() === '確定後振替';
    });
    if (!rules.length) {
      throw new Error('「月間・例外ルール」に使用=○の「確定後振替」がありません。');
    }

    const applied = [];
    rules.forEach(function (r) {
      const staff = findStaffByRef_(ctx, r['対象']);
      if (!staff) {
        throw new Error('確定後振替の対象職員が見つかりません：' + r['対象']);
      }
      const offDate = normalizeDate_(r['開始日']);
      const onDate = normalizeDate_(r['終了日']);
      if (!offDate || !onDate) {
        throw new Error('確定後振替は開始日（休みにする日）と終了日（出勤する日）が必要です。');
      }
      assertDateInTargetMonth_(ctx, offDate, '休みにする日');
      assertDateInTargetMonth_(ctx, onDate, '出勤する日');

      addGenericLock_(ctx, offDate, staff.id, '休み固定', '確定後振替');
      const unit = normalizeUnit_(r['値1']);
      const onType = unit === '第1' ? '第1固定' : unit === '第2' ? '第2固定' : '出勤固定';
      addGenericLock_(ctx, onDate, staff.id, onType, '確定後振替');
      applied.push(staff.name + ' ' + offDate.getDate() + '日→' + onDate.getDate() + '日');
      appendActualChangeLog_(ctx.ss, ctx, staff, offDate, onDate, unit, String(r['メモ'] || ''));
    });

    generateShift_({mode:'reoptimize', scope:{type:'all'}});
    initializeActualSheet_(ctx, readCurrentPlan_(ctx));
    writeOperationStatus_(ctx.ss, '🔵 確定後変更あり', '変更比較と実績ログを確認', applied.join(' / '));
    appendLog_('INFO', '確定後振替', '', applied.join(' / '));
  });
}

/**
 * 現在ファイルをコピーし、対象月を1か月進めて月次入力だけ初期化する。
 * 施設・職員・働き方・勤務パターン・施設/個人/ペアルールは引き継ぐ。
 */
function shiftCreateNextMonthFile() {
  runSafely_('次月ファイル作成', function () {
    const ctx = loadContext_();
    const nextMonth = new Date(
      ctx.targetMonth.getFullYear(),
      ctx.targetMonth.getMonth() + 1,
      1
    );
    const label = Utilities.formatDate(nextMonth, ctx.ss.getSpreadsheetTimeZone(), 'yyyy年MM月');
    const file = DriveApp.getFileById(ctx.ss.getId());
    const copy = file.makeCopy(ctx.ss.getName() + '_' + label);
    const next = SpreadsheetApp.openById(copy.getId());

    const setup = getSheetOrThrow_(next, SHIFT_APP.SHEETS.SETUP);
    setLabelValue_(setup, '最初に作る対象月', nextMonth);

    const request = next.getSheetByName(SHIFT_APP.SHEETS.REQUEST);
    if (request) request.getRange('C6:AG20').clearContent();

    const usage = next.getSheetByName(SHIFT_APP.SHEETS.USAGE);
    if (usage) {
      usage.getRange('C5:D35').clearContent();
      usage.getRange('J5:K35').clearContent();
    }

    const lock = next.getSheetByName(SHIFT_APP.SHEETS.LOCK);
    if (lock) {
      lock.getRange('A5:C80').clearContent();
      lock.getRange('E5:G80').clearContent();
    }

    const ex = next.getSheetByName(SHIFT_APP.SHEETS.EXCEPTION_RULE);
    if (ex) clearExpiredMonthlyExceptionRules_(ex, ctx.targetMonth);

    const auto = next.getSheetByName(SHIFT_APP.SHEETS.AUTO);
    if (auto) {
      auto.getRange(SHIFT_APP.RANGE.WORK_PLAN).clearContent();
      auto.getRange(SHIFT_APP.RANGE.UNIT_PLAN).clearContent();
    }

    const backup = next.getSheetByName(SHIFT_APP.SHEETS.BACKUP);
    if (backup) backup.clearContents();
    const log = next.getSheetByName(SHIFT_APP.SHEETS.LOG);
    if (log) {
      log.clearContents();
      ensureLogHeader_(log);
    }

    const oldSnaps = next.getSheets().filter(function (sh) {
      return sh.getName().indexOf('_確定_') === 0;
    });
    oldSnaps.forEach(function (sh) { next.deleteSheet(sh); });
    const actual = next.getSheetByName(SHIFT_APP.SHEETS.ACTUAL);
    if (actual) actual.getRange('A5:M600').clearContent();
    const reason = next.getSheetByName(SHIFT_APP.SHEETS.REASON);
    if (reason) reason.getRange('A5:H600').clearContent();
    writeOperationStatus_(next, '⚪ 翌月準備済', '希望休・利用予定を入力', label);

    showAlert_('次月ファイルを作成しました', label + '\n\n' + copy.getUrl());
  });
}

/**
 * 本番ブックを変更せず試せるコピーを作る。
 */
function shiftCreateTrialCopy() {
  runSafely_('試算用コピー作成', function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
    const file = DriveApp.getFileById(ss.getId());
    const copy = file.makeCopy(ss.getName() + '_試算_' + stamp);
    showAlert_(
      '試算用コピーを作成しました',
      '本番は変更していません。\n\n試算コピー：\n' + copy.getUrl()
    );
  });
}

/**
 * 指定日のみ再生成する。
 * 他の日は現在のシフトを固定したまま処理する。
 */
function shiftRegenerateDay() {
  runSafely_('指定日再生成', function () {
    const ctx = loadContext_();
    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
      '指定日のみ再生成',
      '再生成する日を 1～' + ctx.daysInMonth + ' の数字で入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const day = Number(response.getResponseText());
    if (!Number.isInteger(day) || day < 1 || day > ctx.daysInMonth) {
      throw new Error(
        '入力エラー：日付は 1～' + ctx.daysInMonth + ' の整数で入力してください。'
      );
    }

    generateShift_({
      mode: 'partial',
      scope: { type: 'day', day: day }
    });
  });
}


/**
 * 指定職員のみ再生成する。
 * 他の職員は現在のシフトを固定する。
 */
function shiftRegenerateStaff() {
  runSafely_('指定職員再生成', function () {
    const ctx = loadContext_();
    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
      '指定職員のみ再生成',
      '職員IDを入力してください。\n例：ST003',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const staffId = String(response.getResponseText() || '').trim();
    const staff = ctx.staff.find(function (x) {
      return x.id === staffId && x.active;
    });

    if (!staff) {
      throw new Error(
        '入力エラー：在籍中の職員ID「' + staffId + '」が見つかりません。'
      );
    }

    generateShift_({
      mode: 'partial',
      scope: { type: 'staff', staffId: staffId }
    });
  });
}


/**
 * 指定単位のみ再生成する。
 * 他単位に配置されている勤務は原則固定する。
 */
function shiftRegenerateUnit() {
  runSafely_('指定単位再生成', function () {
    const ctx = loadContext_();
    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
      '指定単位のみ再生成',
      '「第1」または「第2」と入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const unit = normalizeUnit_(response.getResponseText());

    if (unit !== '第1' && unit !== '第2') {
      throw new Error(
        '入力エラー：「第1」または「第2」と入力してください。'
      );
    }

    if (unit === '第2' && ctx.unitCount < 2) {
      throw new Error(
        '設定エラー：この施設は1単位設定のため、第2単位は再生成できません。'
      );
    }

    generateShift_({
      mode: 'partial',
      scope: { type: 'unit', unit: unit }
    });
  });
}


/**
 * 当日欠勤を配置ロックへ登録し、その日だけ再生成する。
 */
function shiftRegisterAbsence() {
  runSafely_('当日欠勤反映', function () {
    const ctx = loadContext_();
    const ui = SpreadsheetApp.getUi();

    const dayResponse = ui.prompt(
      '当日欠勤を反映',
      '欠勤する日を 1～' + ctx.daysInMonth + ' の数字で入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

    if (dayResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const day = Number(dayResponse.getResponseText());
    if (!Number.isInteger(day) || day < 1 || day > ctx.daysInMonth) {
      throw new Error(
        '入力エラー：日付は 1～' + ctx.daysInMonth + ' の整数で入力してください。'
      );
    }

    const staffResponse = ui.prompt(
      '当日欠勤を反映',
      '欠勤する職員IDを入力してください。\n例：ST003',
      ui.ButtonSet.OK_CANCEL
    );

    if (staffResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const staffId = String(staffResponse.getResponseText() || '').trim();
    const staff = ctx.staff.find(function (x) {
      return x.id === staffId && x.active;
    });

    if (!staff) {
      throw new Error(
        '入力エラー：在籍中の職員ID「' + staffId + '」が見つかりません。'
      );
    }

    addAbsenceLock_(ctx, day, staffId);

    generateShift_({
      mode: 'partial',
      scope: { type: 'day', day: day }
    });
  });
}


/**
 * 直前バックアップを復元する。
 */
function shiftRestoreBackup() {
  runSafely_('バックアップ復元', function () {
    restoreBackup_();
  });
}


/* =========================================================
 * 3. 生成本体
 * ========================================================= */

/**
 * 月間シフトを生成する中心処理。
 *
 * @param {Object} options
 *   mode: generate / reoptimize / partial
 *   scope:
 *     {type:'all'}
 *     {type:'day', day:15}
 *     {type:'staff', staffId:'ST003'}
 *     {type:'unit', unit:'第2'}
 */
function generateShift_(options) {
  const ctx = loadContext_();
  const setupCheck = validateInitialSetup_(ctx);

  if (setupCheck.errors.length > 0) {
    writeValidationLog_('生成前チェック', setupCheck);
    showAlert_(
      'シフトを作成できません',
      '初期設定に修正が必要です。\n\n' + formatValidationResult_(setupCheck)
    );
    return;
  }

  if (!ctx.usageEntered) {
    throw new Error(
      '入力エラー：「利用予定」に1日分も利用予定人数が入力されていません。'
    );
  }

  // 実行前に必ずバックアップ。
  backupCurrentState_(
    options.mode === 'partial'
      ? '部分再生成'
      : options.mode === 'reoptimize'
        ? '再最適化'
        : '月間生成'
  );

  const existing = readCurrentPlan_(ctx);
  const state = createInitialState_(ctx, existing, options);

  // 1. 施設休業・不可・有休等のハード条件を反映
  enforceHardAvailability_(ctx, state);

  // 2. 配置ロックを反映
  const lockResult = applyPlacementLocks_(ctx, state);
  if (lockResult.errors.length > 0) {
    throw new Error(
      '配置ロックに矛盾があります。\n\n' +
      lockResult.errors.map(function (x) { return '・' + x; }).join('\n')
    );
  }

  // 3. 契約勤務日数を基準に初期勤務日を決める
  seedEmployeeWorkDays_(ctx, state);

  // 3.5 高度ルールの固定出勤・休み・Unit固定を反映
  applyAdvancedRuleLocks_(ctx, state);

  // 4. 児発管を各単位へ確保
  ensureJihatsukan_(ctx, state);

  // 5. 所属が共通・未設定の出勤者を必要単位へ仮配置
  assignUnassignedScheduledStaff_(ctx, state);

  // 6. 法定配置を不足しないよう補正
  ensureLegalStaffing_(ctx, state);

  // 7. 送迎人数を確保
  ensureTransportStaffing_(ctx, state);

  // 7.5 施設ルールの役割組合せを確保
  ensureRoleCombinationRules_(ctx, state);

  // 8. 加配・専門支援をソフト条件として可能な範囲で補正
  applySoftRoleTargets_(ctx, state);

  // 9. ペアルールを可能な範囲で補正
  repairPairRules_(ctx, state);

  // 10. 契約日数を再調整
  normalizeContractDays_(ctx, state);

  // 11. 最低連休を勤務日の入替で可能な範囲まで確保
  repairMinimumConsecutiveRest_(ctx, state);

  // 12. 契約時間を、勤務日数を変えず勤務日の入替で可能な範囲まで補正
  repairContractHoursByDaySwap_(ctx, state);

  // 13. 最大連勤を補正
  repairMaxConsecutive_(ctx, state);

  // 14. 原則連勤数を可能な範囲で改善
  improvePreferredConsecutive_(ctx, state);

  // 15. 補正後に法定・送迎・役割組合せを再確認
  ensureLegalStaffing_(ctx, state);
  ensureTransportStaffing_(ctx, state);
  ensureRoleCombinationRules_(ctx, state);

  // 16. 内部グリッドへ確定
  writePlan_(ctx, state);
  writePlacementReasons_(ctx, state);

  SpreadsheetApp.flush();

  // 17. 出力後検証
  const finalPlan = readCurrentPlan_(ctx);
  const result = validateCurrentShift_(ctx, finalPlan);

  writeValidationLog_('生成完了', result);

  const suggestions = buildCorrectionSuggestions_(ctx, finalPlan, result, 10);
  writeCorrectionSuggestions_(ctx, suggestions, result);
  writeChangeComparison_(ctx, existing, finalPlan,
    options.mode === 'partial' ? '部分再生成' :
    options.mode === 'reoptimize' ? '再最適化' : '月間生成');

  writeOperationStatus_(
    ctx.ss,
    result.errors.length ? '🔴 要修正' : '🟡 作成済',
    result.errors.length ? '修正候補を確認' : '最終チェック→シフト確定',
    result.errors.length ? result.errors.slice(0,3).join(' / ') : (result.warnings.slice(0,3).join(' / ') || '重大な未解決なし')
  );

  appendLog_(
    'INFO',
    '生成処理完了',
    Utilities.formatDate(
      ctx.targetMonth,
      ctx.ss.getSpreadsheetTimeZone(),
      'yyyy-MM'
    ),
    '検証・修正候補・変更比較まで更新しました。'
  );

  let message = '';
  if (result.errors.length === 0) {
    message += 'シフト生成が完了しました。';
  } else {
    message += 'シフトは生成しましたが、安全条件に未解決項目があります。';
  }

  message += '\n\n' + formatValidationResult_(result);

  if (state.warnings.length > 0) {
    message += '\n\n【生成時の注意】\n' +
      unique_(state.warnings).slice(0, 12).map(function (x) {
        return '・' + x;
      }).join('\n');
  }

  if (suggestions.length > 0) {
    message += '\n\n【修正候補】\n' +
      suggestions.slice(0, 3).map(function (x, i) {
        return (i + 1) + '. ' + x;
      }).join('\n');
  }

  showAlert_('シフト自動作成', message);
}


/* =========================================================
 * 4. データ読込
 * ========================================================= */

/**
 * シート全体の必要情報を読み、GAS内の共通データへ変換する。
 */
function loadContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const setupSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.SETUP);
  const staffSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.STAFF);
  const workStyleSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.WORK_STYLE);
  const patternSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.WORK_PATTERN);
  const placementSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.PLACEMENT);
  const requestSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.REQUEST);
  const usageSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.USAGE);
  const autoSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.AUTO);
  const lockSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.LOCK);

  const setupValues = setupSheet.getRange('A1:D80').getValues();

  const targetMonth = firstDayOfMonth_(
    findLabelValue_(setupValues, '最初に作る対象月') ||
    findLabelValue_(setupValues, '対象年月')
  );

  if (!(targetMonth instanceof Date) || isNaN(targetMonth.getTime())) {
    throw new Error(
      '設定エラー：「初回セットアップ」の対象月を正しく入力してください。'
    );
  }

  const daysInMonth = new Date(
    targetMonth.getFullYear(),
    targetMonth.getMonth() + 1,
    0
  ).getDate();

  const unitCount = numberOrZero_(findLabelValue_(setupValues, '単位数'));
  const preferredConsecutive =
    numberOrZero_(findLabelValue_(setupValues, '原則連勤')) || 3;
  const maxConsecutive =
    numberOrZero_(findLabelValue_(setupValues, '最大連勤')) || 4;
  const requestOffLimit =
    numberOrZero_(findLabelValue_(setupValues, '希望休上限'));
  const municipality =
    String(findLabelValue_(setupValues, '運用自治体') || '').trim();
  const businessHours = {
    weekday: readSetupClockRange_(setupValues, '届出営業時間（放課後）'),
    holiday: readSetupClockRange_(setupValues, '届出営業時間（学校休日）')
  };

  const transportUse =
    String(findLabelValue_(setupValues, '送迎') || '').trim();
  const transportVehicleCount =
    numberOrZero_(findLabelValue_(setupValues, '通常運用車両数'));
  const employeeRideRule =
    String(findLabelValue_(setupValues, '各車両の社員同乗') || '').trim();

  // 職員セットアップ
  const staffValues = staffSheet.getRange('A4:AF19').getValues();
  const staffHeaders = staffValues[0].map(String);
  const staffHeader = buildHeaderMap_(staffHeaders);

  // 自動計算：必要勤務日数・時間等
  const autoMetaValues = autoSheet.getRange(SHIFT_APP.RANGE.AUTO_META).getValues();
  const autoMetaHeaders = autoMetaValues[0].map(String);
  const autoMetaHeader = buildHeaderMap_(autoMetaHeaders);

  // 既存スコア
  const scoreValues = autoSheet.getRange(SHIFT_APP.RANGE.SCORE).getValues();
  const halfDayValues = autoSheet.getRange(SHIFT_APP.RANGE.HALF_DAY).getValues();

  // 希望入力
  const requestValues = requestSheet.getRange('A5:AJ20').getValues();

  const staff = [];

  for (let i = 0; i < SHIFT_APP.MAX_STAFF; i++) {
    const row = staffValues[i + 1] || [];
    const meta = autoMetaValues[i + 1] || [];
    const requestRow = requestValues[i + 1] || [];
    const halfRow = halfDayValues[i] || [];

    const id = stringAt_(row, staffHeader, '職員ID');
    const name = stringAt_(row, staffHeader, '氏名');
    const employment = stringAt_(row, staffHeader, '雇用区分');
    const activeText = stringAt_(row, staffHeader, '在籍');

    const requiredDays = numberAt_(meta, autoMetaHeader, '必要日数');
    const minimumHours = numberAt_(meta, autoMetaHeader, '最低時間');
    const targetHours = numberAt_(meta, autoMetaHeader, '目標時間');

    const preferences = [];
    for (let d = 0; d < SHIFT_APP.MAX_DAYS; d++) {
      preferences.push(String(requestRow[d + 2] || '').trim());
    }

    const specialistType = stringAt_(row, staffHeader, '専門職種（主）') ||
      stringAt_(row, staffHeader, 'その他専門職');
    const isOtherSpecialist = specialistType !== '' &&
      specialistType !== 'なし' && specialistType !== '×' && specialistType !== '要確認';

    staff.push({
      index: i,
      id: id,
      name: name,
      employment: employment,
      active: activeText === '在籍' && id !== '' && name !== '',
      unit: normalizeUnit_(stringAt_(row, staffHeader, '所属')),
      workStyle: stringAt_(row, staffHeader, '働き方'),
      role: stringAt_(row, staffHeader, '役職'),
      sex: stringAt_(row, staffHeader, '性別(配置判定)'),
      isJihatsukanQualified: isYes_(valueAt_(row, staffHeader, '児発管')),
      isChildInstructor: isYes_(valueAt_(row, staffHeader, '児童指導員')),
      isNursery: isYes_(valueAt_(row, staffHeader, '保育士')),
      specialistType: specialistType,
      isOtherSpecialist: isOtherSpecialist,
      directSupport: isYes_(valueAt_(row, staffHeader, '直接支援配置算入')),
      legalEligible: isYes_(valueAt_(row, staffHeader, '児童指導員配置可')),
      specialistEligible: isYes_(valueAt_(row, staffHeader, '専門支援配置可')),
      jihatsukanEligible: isYes_(valueAt_(row, staffHeader, '児発管配置可')),
      addOnCategory: stringAt_(row, staffHeader, '加配資格区分'),
      canDrive: isYes_(valueAt_(row, staffHeader, '運転可')),
      canRide: isYes_(valueAt_(row, staffHeader, '添乗可')),
      canTransport: isYes_(valueAt_(row, staffHeader, '送迎可')),
      fixedWorkWeekdays: stringAt_(row, staffHeader, '固定勤務曜日'),
      fixedDaysOff: stringAt_(row, staffHeader, '固定休日'),
      restrictionMemo: stringAt_(row, staffHeader, '勤務制限・備考'),
      adminRegisteredRole: stringAt_(row, staffHeader, '行政登録職種'),
      adminEmploymentClass: stringAt_(row, staffHeader, '常勤区分'),
      adminExclusivity: stringAt_(row, staffHeader, '専従・兼務'),
      adminConcurrentRole: stringAt_(row, staffHeader, '兼務職種'),
      adminNoticeStatus: stringAt_(row, staffHeader, '届出確認'),
      adminMemo: stringAt_(row, staffHeader, '行政メモ'),
      requiredDays: requiredDays,
      minimumHours: minimumHours,
      targetHours: targetHours,
      preferences: preferences,
      paidLeaveDays: preferences.slice(0, daysInMonth).filter(function (x) {
        return x === '有休';
      }).length,
      scores: (scoreValues[i] || []).map(function (x) {
        const n = Number(x);
        return Number.isFinite(n) ? n : -999999;
      }),
      halfDays: halfRow.map(function (x) {
        const n = Number(x);
        return Number.isFinite(n) ? n : 0;
      }).filter(function (x) {
        return x > 0;
      })
    });
  }

  // 利用予定
  const usageValues = usageSheet.getRange('A4:L35').getValues();
  const usageHeaders = usageValues[0].map(String);
  const usageHeader = buildHeaderMap_(usageHeaders);
  const days = [];
  let usageEntered = false;

  for (let d = 0; d < daysInMonth; d++) {
    const row = usageValues[d + 1] || [];

    const rawUsage1 = valueAt_(row, usageHeader, '第1予定人数');
    const rawUsage2 = valueAt_(row, usageHeader, '第2予定人数');

    if (rawUsage1 !== '' && rawUsage1 !== null) usageEntered = true;
    if (unitCount >= 2 && rawUsage2 !== '' && rawUsage2 !== null) usageEntered = true;

    days.push({
      day: d + 1,
      date: new Date(
        targetMonth.getFullYear(),
        targetMonth.getMonth(),
        d + 1
      ),
      weekday: weekdayJa_(new Date(
        targetMonth.getFullYear(),
        targetMonth.getMonth(),
        d + 1
      )),
      usage1: numberOrZero_(rawUsage1),
      usage2: unitCount >= 2 ? numberOrZero_(rawUsage2) : 0,
      legal1: numberOrZero_(valueAt_(row, usageHeader, '第1法定必要')),
      legal2: unitCount >= 2
        ? numberOrZero_(valueAt_(row, usageHeader, '第2法定必要'))
        : 0,
      state: String(valueAt_(row, usageHeader, '営業状態') || '').trim(),
      dayType: String(valueAt_(row, usageHeader, '日区分（自動）') || '').trim()
    });
  }

  const placementValues = placementSheet.getRange('A1:J100').getValues();
  const placement = parsePlacementSettings_(placementValues, unitCount);

  const lockValues = lockSheet.getRange('A4:H80').getValues();
  const locks = parseLocks_(lockValues, targetMonth, daysInMonth);

  const workStyleValues = workStyleSheet.getRange('A4:N24').getValues();
  const patternValues = patternSheet.getRange('A4:L20').getValues();
  const serviceTimeSheet = ss.getSheetByName(SHIFT_APP.SHEETS.SERVICE_TIME);
  const serviceTimeValues = serviceTimeSheet
    ? serviceTimeSheet.getRange('A3:F10').getDisplayValues()
    : [];

  const rules = loadAdvancedRules_(ss);
  applyAdvancedDayRules_(days, rules, transportVehicleCount, unitCount);
  attachStaffCostRates_(staff, rules);

  const rulePreferred = numberOrZero_(facilityRuleValue_(rules, '原則連勤', preferredConsecutive));
  const ruleMax = numberOrZero_(facilityRuleValue_(rules, '最大連勤', maxConsecutive));

  return {
    ss: ss,
    targetMonth: targetMonth,
    daysInMonth: daysInMonth,
    unitCount: unitCount,
    preferredConsecutive: rulePreferred > 0 ? rulePreferred : preferredConsecutive,
    maxConsecutive: ruleMax > 0 ? ruleMax : maxConsecutive,
    requestOffLimit: requestOffLimit,
    municipality: municipality,
    businessHours: businessHours,
    transportUse: transportUse,
    transportVehicleCount: transportVehicleCount,
    employeeRideRule: employeeRideRule,
    staff: staff,
    days: days,
    usageEntered: usageEntered,
    placement: placement,
    locks: locks,
    workStyleValues: workStyleValues,
    patternValues: patternValues,
    serviceTimeValues: serviceTimeValues,
    rules: rules
  };
}




function isWoodyTokyoMunicipality_(ctx) {
  return ctx && (ctx.municipality === '府中市' || ctx.municipality === '国分寺市');
}

function hasAdminRole_(staff, role) {
  const wanted = String(role || '').trim();
  if (!staff || !wanted) return false;
  return [
    staff.adminRegisteredRole,
    staff.adminConcurrentRole
  ].some(function (x) {
    return String(x || '').trim() === wanted;
  });
}

function isTokyoCoreLegalStaff_(staff) {
  return hasAdminRole_(staff, '児童指導員') ||
    hasAdminRole_(staff, '保育士');
}

function parseClockMinutes_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return NaN;
  const h = Number(m[1]);
  const min = Number(m[2]);
  if (!Number.isFinite(h) || !Number.isFinite(min) || h < 0 || h > 24 || min < 0 || min > 59) return NaN;
  return h * 60 + min;
}

function readSetupClockRange_(values, label) {
  for (let r = 0; r < values.length; r++) {
    const row = values[r] || [];
    for (let c = 0; c < row.length - 2; c++) {
      if (String(row[c] || '').trim() !== label) continue;
      const start = parseClockMinutes_(row[c + 1]);
      const end = parseClockMinutes_(row[c + 2]);
      if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) return null;
      return {start: start, end: end};
    }
  }
  return null;
}

function getServiceWindow_(ctx, label) {
  const rows = ctx.serviceTimeValues || [];
  if (!rows.length) return null;
  const header = buildHeaderMap_((rows[0] || []).map(String));
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (String(valueAt_(row, header, '区分') || '').trim() !== label) continue;
    const start = parseClockMinutes_(valueAt_(row, header, '開始'));
    const end = parseClockMinutes_(valueAt_(row, header, '終了'));
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    return {start: start, end: end};
  }
  return null;
}

function getPatternWindow_(ctx, patternName) {
  const rows = ctx.patternValues || [];
  if (!rows.length) return null;
  const header = buildHeaderMap_((rows[0] || []).map(String));
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const name = String(
      valueAt_(row, header, '勤務パターン名') ||
      valueAt_(row, header, 'パターン名') ||
      ''
    ).trim();
    if (name !== patternName) continue;
    if (String(valueAt_(row, header, '使用') || '').trim() !== '○') return null;
    if (String(valueAt_(row, header, '配置算入') || '').trim() !== '○') return null;
    const start = parseClockMinutes_(valueAt_(row, header, '開始'));
    const end = parseClockMinutes_(valueAt_(row, header, '終了'));
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    return {start: start, end: end};
  }
  return null;
}

function windowCovers_(pattern, service) {
  return !!pattern && !!service &&
    pattern.start <= service.start &&
    pattern.end >= service.end;
}

function validateTokyoAdminCompliance_(ctx) {
  const errors = [];
  const warnings = [];

  if (!isWoodyTokyoMunicipality_(ctx)) {
    errors.push('運用自治体は「府中市」または「国分寺市」を選択してください。');
    return {errors: errors, warnings: warnings};
  }

  const active = ctx.staff.filter(function (staff) { return staff.active; });

  active.forEach(function (staff) {
    const label = staff.id + ' ' + staff.name;
    if (!staff.adminRegisteredRole) {
      errors.push(label + '：行政登録職種が未入力です。');
    }
    if (!staff.adminEmploymentClass) {
      errors.push(label + '：常勤区分が未入力です。');
    }
    if (!staff.adminExclusivity) {
      errors.push(label + '：専従・兼務が未入力です。');
    }
    if (staff.adminNoticeStatus !== '確認済') {
      errors.push(label + '：行政届出の確認状態が「確認済」ではありません。');
    }
    if (staff.legalEligible && !isTokyoCoreLegalStaff_(staff)) {
      errors.push(
        label + '：「児童指導員配置可」が○ですが、行政登録職種が児童指導員または保育士として確認できません。'
      );
    }
  });

  for (let u = 1; u <= ctx.unitCount; u++) {
    const unit = '第' + u;
    const fullTimeCore = active.some(function (staff) {
      return staff.unit === unit &&
        staff.adminEmploymentClass === '常勤' &&
        isTokyoCoreLegalStaff_(staff) &&
        staff.adminNoticeStatus === '確認済';
    });
    if (!fullTimeCore) {
      errors.push(unit + '：常勤の児童指導員または保育士を1名以上確認できません。');
    }
  }

  const hasManager = active.some(function (staff) {
    return hasAdminRole_(staff, '管理者') &&
      staff.adminNoticeStatus === '確認済';
  });
  if (!hasManager) {
    errors.push('管理者の行政登録を確認できません。');
  }

  const configuredJihatsu = [];
  for (let u = 1; u <= ctx.unitCount; u++) {
    const unit = '第' + u;
    const cfg = ctx.placement[unit] && ctx.placement[unit].jihatsukan;
    if (!cfg) continue;
    [cfg.main, cfg.sub].filter(Boolean).forEach(function (ref) {
      const staff = findStaffByRef_(ctx, ref);
      if (staff && configuredJihatsu.indexOf(staff) < 0) configuredJihatsu.push(staff);
    });
  }

  const hasDedicatedFullTimeJihatsu = configuredJihatsu.some(function (staff) {
    const exclusivityOk =
      staff.adminExclusivity === '専従' ||
      (staff.adminExclusivity === '兼務' && hasAdminRole_(staff, '管理者'));
    return staff.active &&
      hasAdminRole_(staff, '児発管') &&
      staff.adminEmploymentClass === '常勤' &&
      exclusivityOk &&
      staff.adminNoticeStatus === '確認済';
  });
  if (!hasDedicatedFullTimeJihatsu) {
    errors.push('専任かつ常勤の児発管（同一事業所の管理者兼務可）を1名以上確認できません。');
  }

  configuredJihatsu.forEach(function (staff) {
    if (!hasAdminRole_(staff, '児発管')) {
      errors.push(staff.name + '：児発管担当ですが行政登録職種に児発管がありません。');
    }
  });

  const weekdayBusiness = ctx.businessHours && ctx.businessHours.weekday;
  const holidayBusiness = ctx.businessHours && ctx.businessHours.holiday;
  const weekdayPattern = getPatternWindow_(ctx, '平日通常');
  const holidayPattern = getPatternWindow_(ctx, '学校休日通常');

  if (!weekdayBusiness) {
    errors.push('届出営業時間（放課後）が未入力または不正です。');
  } else if (!windowCovers_(weekdayPattern, weekdayBusiness)) {
    errors.push('平日通常の配置算入時間が、届出営業時間（放課後）全体をカバーしていません。');
  }

  if (!holidayBusiness) {
    errors.push('届出営業時間（学校休日）が未入力または不正です。');
  } else if (!windowCovers_(holidayPattern, holidayBusiness)) {
    errors.push('学校休日通常の配置算入時間が、届出営業時間（学校休日）全体をカバーしていません。');
  }

  const weekdayService = getServiceWindow_(ctx, '放課後');
  if (weekdayBusiness && weekdayService && !windowCovers_(weekdayBusiness, weekdayService)) {
    errors.push('届出営業時間（放課後）が標準サービス提供時間をカバーしていません。');
  }

  ['平日学休日', '学休日・代替'].forEach(function (label) {
    const service = getServiceWindow_(ctx, label);
    if (holidayBusiness && service && !windowCovers_(holidayBusiness, service)) {
      errors.push('届出営業時間（学校休日）が「' + label + '」の標準サービス提供時間をカバーしていません。');
    }
  });

  if (ctx.municipality === '府中市') {
    warnings.push('府中市で新規開設する場合は、市への事業所開設相談と東京都側の指定申請手続を別途確認してください。');
  } else if (ctx.municipality === '国分寺市') {
    warnings.push('国分寺市の指導検査は東京都の基準を準用するため、東京都基準と届出内容の一致を確認してください。');
  }

  return {
    errors: unique_(errors),
    warnings: unique_(warnings)
  };
}

/* =========================================================
 * 4.5 高度ルール（施設・個人・ペア・月間例外）
 * ========================================================= */

function loadAdvancedRules_(ss) {
  return {
    facilityRules: readRuleObjects_(ss.getSheetByName(SHIFT_APP.SHEETS.FACILITY_RULE), 'A4:L120'),
    personRules: readRuleObjects_(ss.getSheetByName(SHIFT_APP.SHEETS.PERSON_RULE), 'A4:L220'),
    pairRules: readRuleObjects_(ss.getSheetByName(SHIFT_APP.SHEETS.PAIR_RULE), 'A4:L160'),
    exceptionRules: readRuleObjects_(ss.getSheetByName(SHIFT_APP.SHEETS.EXCEPTION_RULE), 'A4:L220'),
    costMap: readCostMap_(ss)
  };
}

function readRuleObjects_(sheet, rangeA1) {
  if (!sheet) return [];
  const values = sheet.getRange(rangeA1).getValues();
  if (!values.length) return [];
  const headers = (values[0] || []).map(function (x) { return String(x || '').trim(); });
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    if (row.every(function (v) { return String(v || '').trim() === ''; })) continue;
    const obj = {_row: r + 4};
    headers.forEach(function (h, c) { if (h) obj[h] = row[c]; });
    out.push(obj);
  }
  return out;
}

function readCostMap_(ss) {
  const sh = ss.getSheetByName(SHIFT_APP.SHEETS.COST);
  if (!sh) return {};
  const values = sh.getRange('A4:I19').getValues();
  const headers = (values[0] || []).map(String);
  const h = buildHeaderMap_(headers);
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const id = String(valueAt_(row, h, '職員ID') || '').trim();
    const name = String(valueAt_(row, h, '氏名') || '').trim();
    if (!id && !name) continue;
    const item = {
      method: String(valueAt_(row, h, '計算方式') || '').trim(),
      fixed: numberOrZero_(valueAt_(row, h, '月額固定費')),
      hourly: numberOrZero_(valueAt_(row, h, '時間単価'))
    };
    if (id) map[id] = item;
    if (name) map[name] = item;
  }
  return map;
}

function attachStaffCostRates_(staffList, rules) {
  const map = rules && rules.costMap || {};
  staffList.forEach(function (staff) {
    const item = map[staff.id] || map[staff.name] || {};
    let hourly = numberOrZero_(item.hourly);
    if (!(hourly > 0) && numberOrZero_(item.fixed) > 0 && staff.targetHours > 0) {
      hourly = numberOrZero_(item.fixed) / staff.targetHours;
    }
    staff.costPerHour = hourly > 0 ? hourly : 0;
  });
}

function isRuleEnabled_(rule) {
  return String(rule && rule['使用'] || '').trim() === '○';
}

function facilityRule_(rules, name) {
  const rows = rules && rules.facilityRules || [];
  return rows.filter(function (r) {
    return isRuleEnabled_(r) && String(r['ルール名'] || '').trim() === name;
  }).sort(function (a, b) {
    return numberOrZero_(b['優先度']) - numberOrZero_(a['優先度']);
  })[0] || null;
}

function facilityRuleValue_(rules, name, fallback) {
  const rule = facilityRule_(rules, name);
  return rule ? rule['値'] : fallback;
}

function facilityRuleEnabled_(rules, name) {
  const rule = facilityRule_(rules, name);
  if (!rule) return false;
  const v = String(rule['値'] || '').trim();
  return v !== '無効' && v !== 'OFF' && v !== '0' && v !== 'しない';
}

function facilityRulePriority_(rules, name, fallback) {
  const rule = facilityRule_(rules, name);
  const n = rule ? numberOrZero_(rule['優先度']) : 0;
  return n > 0 ? n : fallback;
}

function facilityRulePriorityForDay_(ctx, name, dayIndex, fallback) {
  let priority = facilityRulePriority_(ctx.rules, name, fallback);
  const base = facilityRule_(ctx.rules, name);
  const baseId = base ? String(base['ルールID'] || '').trim() : '';
  getExceptionRules_(ctx, '一時的優先度変更', dayIndex).forEach(function (r) {
    const target = String(r['対象'] || '').trim();
    if (target !== name && (!baseId || target !== baseId)) return;
    const n = numberOrZero_(r['値1']);
    if (n >= 0 && n <= 100) priority = n;
  });
  return priority;
}

function ruleStrengthMultiplier_(rule) {
  const s = String(rule && rule['強度'] || '').trim();
  if (s === '絶対') return 1.5;
  if (s === 'おまかせ') return 0.5;
  return 1;
}

function ruleDateMatches_(ctx, rule, dayIndex) {
  const day = ctx.days[dayIndex];
  if (!day) return false;
  const date = day.date;
  const start = normalizeDate_(rule['開始日']);
  const end = normalizeDate_(rule['終了日']);
  if (start && date.getTime() < start.getTime()) return false;
  if (end && date.getTime() > end.getTime()) return false;
  const selector = String(rule['曜日/日区分'] || rule['曜日'] || '全日').trim();
  return daySelectorMatches_(day, selector);
}

function daySelectorMatches_(day, selector) {
  const s = String(selector || '全日').trim();
  if (!s || s === '全日') return true;
  if (['月','火','水','木','金','土','日'].indexOf(s) >= 0) return day.weekday === s;
  if (s === '平日') return ['月','火','水','木','金'].indexOf(day.weekday) >= 0;
  if (s === '土日') return day.weekday === '土' || day.weekday === '日';
  if (s === '学校休日') return String(day.dayType || '').indexOf('学校休日') >= 0;
  if (s === '祝日') return String(day.dayType || '').indexOf('祝') >= 0;
  return String(day.dayType || '') === s;
}

function isWeekendLikeDay_(day) {
  if (!day) return false;
  return day.weekday === '土' || day.weekday === '日' ||
    String(day.dayType || '').indexOf('祝') >= 0;
}

function staffRuleMatches_(rule, staff) {
  const target = String(rule && rule['対象職員'] || '').trim();
  return !!target && (target === staff.id || target === staff.name);
}

function getPersonRules_(ctx, staff, type, dayIndex) {
  const rows = ctx.rules && ctx.rules.personRules || [];
  return rows.filter(function (r) {
    if (!isRuleEnabled_(r) || !staffRuleMatches_(r, staff)) return false;
    if (type && String(r['ルール種別'] || '').trim() !== type) return false;
    return dayIndex === undefined || dayIndex === null || ruleDateMatches_(ctx, r, dayIndex);
  }).sort(function (a, b) {
    return numberOrZero_(b['優先度']) - numberOrZero_(a['優先度']);
  });
}

function getExceptionRules_(ctx, type, dayIndex) {
  const rows = ctx.rules && ctx.rules.exceptionRules || [];
  return rows.filter(function (r) {
    if (!isRuleEnabled_(r)) return false;
    if (type && String(r['ルール種別'] || '').trim() !== type) return false;
    return dayIndex === undefined || dayIndex === null || ruleDateMatches_(ctx, r, dayIndex);
  }).sort(function (a, b) {
    return numberOrZero_(b['優先度']) - numberOrZero_(a['優先度']);
  });
}

function advancedHardOffReason_(ctx, staff, dayIndex) {
  const personal = getPersonRules_(ctx, staff, '休み固定', dayIndex).filter(function (r) {
    return String(r['強度'] || '').trim() === '絶対';
  });
  if (personal.length) return '個人ルール休み固定';

  const ex = getExceptionRules_(ctx, '休み固定', dayIndex).filter(function (r) {
    const target = String(r['対象'] || '').trim();
    return (target === staff.id || target === staff.name) &&
      String(r['強度'] || '').trim() === '絶対';
  });
  if (ex.length) return '月間例外休み固定';
  return '';
}

function hasAdvancedFixedWork_(ctx, staff, dayIndex) {
  if (getPersonRules_(ctx, staff, '出勤固定', dayIndex).some(function (r) {
    return String(r['強度'] || '').trim() === '絶対';
  })) return true;
  return getExceptionRules_(ctx, '出勤固定', dayIndex).some(function (r) {
    const target = String(r['対象'] || '').trim();
    return (target === staff.id || target === staff.name) &&
      String(r['強度'] || '').trim() === '絶対';
  });
}

function isAutoAdjustBlocked_(ctx, staff, dayIndex) {
  if (getPersonRules_(ctx, staff, '自動調整禁止', dayIndex).length) return true;
  if (getPersonRules_(ctx, staff, null, dayIndex).some(function (r) {
    return String(r['自動調整'] || '').trim() === '手動のみ';
  })) return true;
  return getExceptionRules_(ctx, '自動調整停止', dayIndex).some(function (r) {
    const target = String(r['対象'] || '').trim();
    return !target || target === '全体' || target === staff.id || target === staff.name;
  });
}

function applyAdvancedRuleLocks_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (state.immutable[staff.index][d]) continue;
      if (advancedHardOffReason_(ctx, staff, d)) {
        state.work[staff.index][d] = 0;
        state.unit[staff.index][d] = '';
        state.immutable[staff.index][d] = true;
        state.lockType[staff.index][d] = '高度ルール休み固定';
        continue;
      }
      if (hasAdvancedFixedWork_(ctx, staff, d)) {
        if (hardUnavailableReason_(ctx, staff, d)) {
          state.errors.push((d + 1) + '日 ' + staff.name + '：出勤固定と勤務不可条件が矛盾します。');
          continue;
        }
        state.work[staff.index][d] = 1;
        const unit = getAdvancedFixedUnit_(ctx, staff, d);
        if (unit) setAssignedUnit_(ctx, state, staff.index, d, unit);
        state.immutable[staff.index][d] = true;
        state.lockType[staff.index][d] = '高度ルール出勤固定';
      } else {
        const unit = getAdvancedFixedUnit_(ctx, staff, d);
        if (unit && state.work[staff.index][d] === 1) {
          setAssignedUnit_(ctx, state, staff.index, d, unit);
          state.immutable[staff.index][d] = true;
          state.lockType[staff.index][d] = '高度ルールUnit固定';
        }
      }
    }
  });
}

function getAdvancedFixedUnit_(ctx, staff, dayIndex) {
  const ex = getExceptionRules_(ctx, 'Unit固定', dayIndex).find(function (r) {
    const target = String(r['対象'] || '').trim();
    return (target === staff.id || target === staff.name) &&
      String(r['強度'] || '').trim() === '絶対';
  });
  if (ex) {
    const u = normalizeUnit_(ex['値1']);
    if (u === '第1' || u === '第2') return u;
  }
  const fixed = getPersonRules_(ctx, staff, 'Unit優先', dayIndex).find(function (r) {
    return String(r['強度'] || '').trim() === '絶対';
  });
  if (fixed) {
    const u = normalizeUnit_(fixed['値1']);
    if (u === '第1' || u === '第2') return u;
  }
  return '';
}

function personRuleDayScore_(ctx, staff, dayIndex) {
  let score = 0;
  getPersonRules_(ctx, staff, null, dayIndex).forEach(function (r) {
    const type = String(r['ルール種別'] || '').trim();
    const priority = numberOrZero_(r['優先度']) || 50;
    const w = priority * 100 * ruleStrengthMultiplier_(r);
    if (type === '出勤希望' || (type === '出勤固定' && String(r['強度'] || '').trim() !== '絶対')) score += w;
    if (type === '休み希望' || (type === '休み固定' && String(r['強度'] || '').trim() !== '絶対')) score -= w;
    if (type === '土日祝ウェイト' && isWeekendLikeDay_(ctx.days[dayIndex])) {
      const v = Number(r['値1']);
      score += (Number.isFinite(v) ? v : 0) * priority * 10;
    }
  });
  return score;
}

function exceptionRuleDayScore_(ctx, staff, dayIndex) {
  let score = 0;
  const rows = ctx.rules && ctx.rules.exceptionRules || [];
  rows.forEach(function (r) {
    if (!isRuleEnabled_(r) || !ruleDateMatches_(ctx, r, dayIndex)) return;
    const target = String(r['対象'] || '').trim();
    if (target !== staff.id && target !== staff.name) return;
    if (String(r['強度'] || '').trim() === '絶対') return;
    const type = String(r['ルール種別'] || '').trim();
    const priority = numberOrZero_(r['優先度']) || 50;
    const w = priority * 100 * ruleStrengthMultiplier_(r);
    if (type === '出勤固定') score += w;
    if (type === '休み固定') score -= w;
  });
  return score;
}

function pairRuleDayScore_(ctx, staff, dayIndex, state) {
  if (!state) return 0;
  let score = 0;
  const rows = ctx.rules && ctx.rules.pairRules || [];
  rows.forEach(function (r) {
    if (!isRuleEnabled_(r) || !ruleDateMatches_(ctx, r, dayIndex)) return;
    const a = String(r['職員A'] || '').trim();
    const b = String(r['職員B'] || '').trim();
    if (a !== staff.name && a !== staff.id && b !== staff.name && b !== staff.id) return;
    const otherRef = (a === staff.name || a === staff.id) ? b : a;
    const other = findStaffByRef_(ctx, otherRef);
    if (!other) return;
    const otherWorks = state.work[other.index][dayIndex] === 1;
    const type = String(r['ルール'] || '').trim();
    const priority = numberOrZero_(r['優先度']) || 50;
    const w = priority * 100 * ruleStrengthMultiplier_(r);
    if (type === '合わせる') score += otherWorks ? w : -w;
    if (type === 'どちらか一方' || type === '合わせない') score += otherWorks ? -w : w;
    if (type === '同時休み優先' && !otherWorks) score -= w;
    if (type === '同時休み禁止' && !otherWorks) score += w;
  });
  return score;
}

function countWeekendLikeWork_(ctx, state, staffIndex) {
  let n = 0;
  for (let d = 0; d < ctx.daysInMonth; d++) {
    if (state.work[staffIndex][d] === 1 && isWeekendLikeDay_(ctx.days[d])) n++;
  }
  return n;
}

function getStaffMaxConsecutive_(ctx, staff, dayIndex) {
  const rules = getPersonRules_(ctx, staff, '最大連勤', dayIndex);
  let cap = ctx.maxConsecutive;
  rules.forEach(function (r) {
    const n = numberOrZero_(r['値1']);
    if (n > 0) cap = Math.min(cap, n);
  });
  return cap;
}

function getStaffPreferredConsecutive_(ctx, staff) {
  return Math.min(ctx.preferredConsecutive, getStaffMaxConsecutive_(ctx, staff));
}

function getPersonPatternRule_(ctx, staff, dayIndex) {
  const rule = getPersonRules_(ctx, staff, '勤務パターン指定', dayIndex)[0];
  return rule ? String(rule['値1'] || '').trim() : '';
}

function isUnitAllowedForStaff_(ctx, staff, dayIndex, unit) {
  if (unit !== '第1' && unit !== '第2') return false;
  if (unit === '第2' && ctx.unitCount < 2) return false;
  const home = normalizeUnit_(staff.unit);
  const helpRule = facilityRule_(ctx.rules, 'Unitヘルプ許可');
  if (helpRule && !facilityRuleEnabled_(ctx.rules, 'Unitヘルプ許可') &&
      (home === '第1' || home === '第2') && home !== unit) return false;

  const denied = getPersonRules_(ctx, staff, 'Unit禁止', dayIndex).some(function (r) {
    const target = normalizeUnit_(r['値1']);
    return target === unit && String(r['強度'] || '').trim() === '絶対';
  });
  if (denied) return false;

  const helpDenied = getPersonRules_(ctx, staff, 'ヘルプ禁止', dayIndex).some(function (r) {
    return String(r['強度'] || '').trim() === '絶対';
  });
  if (helpDenied && (home === '第1' || home === '第2') && home !== unit) return false;
  return true;
}

function unitPreferenceScore_(ctx, staff, dayIndex, unit, state) {
  if (!isUnitAllowedForStaff_(ctx, staff, dayIndex, unit)) return -999999;
  let score = 0;
  const home = normalizeUnit_(staff.unit);
  if (home === unit) score += 1000;
  getPersonRules_(ctx, staff, null, dayIndex).forEach(function (r) {
    const type = String(r['ルール種別'] || '').trim();
    const target = normalizeUnit_(r['値1']);
    const w = (numberOrZero_(r['優先度']) || 50) * 20 * ruleStrengthMultiplier_(r);
    if ((type === 'Unit優先' || type === 'ヘルプ優先') && target === unit) score += w;
    if (type === 'Unit禁止' && target === unit) score -= w;
    if (type === 'ヘルプ禁止' && home && home !== unit) score -= w;
  });
  getExceptionRules_(ctx, 'Unit固定', dayIndex).forEach(function (r) {
    const targetStaff = String(r['対象'] || '').trim();
    if (targetStaff !== staff.id && targetStaff !== staff.name) return;
    if (String(r['強度'] || '').trim() === '絶対') return;
    const targetUnit = normalizeUnit_(r['値1']);
    if (targetUnit === unit) {
      score += (numberOrZero_(r['優先度']) || 50) * 20 * ruleStrengthMultiplier_(r);
    }
  });

  if (state) {
    const rows = ctx.rules && ctx.rules.pairRules || [];
    rows.forEach(function (r) {
      if (!isRuleEnabled_(r) || !ruleDateMatches_(ctx, r, dayIndex)) return;
      const a = String(r['職員A'] || '').trim();
      const b = String(r['職員B'] || '').trim();
      if (a !== staff.name && a !== staff.id && b !== staff.name && b !== staff.id) return;
      const other = findStaffByRef_(ctx, (a === staff.name || a === staff.id) ? b : a);
      if (!other || state.work[other.index][dayIndex] !== 1) return;
      const otherUnit = getAssignedUnit_(ctx, state, other.index, dayIndex);
      const w = (numberOrZero_(r['優先度']) || 50) * 20 * ruleStrengthMultiplier_(r);
      if (String(r['ルール'] || '').trim() === '同一Unit優先') score += otherUnit === unit ? w : -w;
      if (String(r['ルール'] || '').trim() === '別Unit優先') score += otherUnit && otherUnit !== unit ? w : -w;
    });
  }
  return score;
}

function choosePreferredUnitForStaff_(ctx, state, staff, dayIndex, fallback) {
  const candidates = ['第1'];
  if (ctx.unitCount >= 2) candidates.push('第2');
  let best = '';
  let bestScore = -Infinity;
  candidates.forEach(function (unit) {
    if (!isUnitAllowedForStaff_(ctx, staff, dayIndex, unit)) return;
    let score = unitPreferenceScore_(ctx, staff, dayIndex, unit, state);
    if (unit === fallback) score += 500;
    const needed = requiredDirectSupportCount_(ctx, ctx.days[dayIndex], unit);
    const deficit = Math.max(0, needed - countLegalStaff_(ctx, state, dayIndex, unit));
    score += deficit * 2000;
    if (score > bestScore) { bestScore = score; best = unit; }
  });
  return best || (isUnitAllowedForStaff_(ctx, staff, dayIndex, fallback) ? fallback : '');
}

function requiredDirectSupportCount_(ctx, day, unit) {
  let needed = unit === '第1' ? numberOrZero_(day.legal1) : numberOrZero_(day.legal2);
  if (day.state === '休日') return 0;
  if (needed <= 0) return needed;
  const rule = facilityRule_(ctx.rules, '最低現場余裕人数');
  if (!rule) return needed;
  const target = String(rule['対象'] || '').trim();
  if (target && target !== '全体' && normalizeUnit_(target) !== unit) return needed;
  const extra = Math.max(0, numberOrZero_(rule['値']));
  return needed + extra;
}

function getTransportNeedForDay_(ctx, day) {
  if (!day || day.state === '休日') return 0;
  if (day.transportVehicleCount === 0) return 0;
  if (day.transportVehicleCount !== undefined && day.transportVehicleCount !== null && day.transportVehicleCount !== '') {
    return Math.max(0, numberOrZero_(day.transportVehicleCount));
  }
  return Math.max(0, numberOrZero_(ctx.transportVehicleCount));
}

function legalRequiredFromUsage_(usage) {
  const n = Math.max(0, numberOrZero_(usage));
  if (n <= 0) return 0;
  return n <= 10 ? 2 : 2 + Math.ceil((n - 10) / 5);
}

function applyAdvancedDayRules_(days, rules, defaultVehicleCount, unitCount) {
  days.forEach(function (day) { day.transportVehicleCount = defaultVehicleCount; });
  const rows = rules && rules.exceptionRules || [];
  rows.forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    days.forEach(function (day) {
      const fakeCtx = {days:[day]};
      if (!ruleDateMatches_(fakeCtx, r, 0)) return;
      const type = String(r['ルール種別'] || '').trim();
      const target = String(r['対象'] || '').trim();
      if (type === '臨時休業') {
        if (!target || target === '全体') {
          day.state = '休日';
          day.legal1 = 0;
          day.legal2 = 0;
        }
        if (normalizeUnit_(target) === '第1') { day.usage1 = 0; day.legal1 = 0; }
        if (normalizeUnit_(target) === '第2' && unitCount >= 2) { day.usage2 = 0; day.legal2 = 0; }
      }
      if (type === '臨時営業' && (!target || target === '全体')) {
        day.state = '営業';
        day.legal1 = legalRequiredFromUsage_(day.usage1);
        day.legal2 = unitCount >= 2 ? legalRequiredFromUsage_(day.usage2) : 0;
      }
      if (type === '祝日指定') day.dayType = '祝日';
      if (type === '最低人数加算') {
        const add = Math.max(0, numberOrZero_(r['値1']));
        if (!target || target === '全体' || normalizeUnit_(target) === '第1') day.legal1 += add;
        if (unitCount >= 2 && (!target || target === '全体' || normalizeUnit_(target) === '第2')) day.legal2 += add;
      }
      if (type === '送迎台数変更') day.transportVehicleCount = Math.max(0, numberOrZero_(r['値1']));
    });
  });
}

function parseRoleCombinationRule_(rule) {
  const text = String(rule['値'] || '').trim();
  if (!text || text.indexOf('+') < 0) return null;
  const parts = text.split(':');
  const caps = parts[0].split('+').map(function (x) { return x.trim(); }).filter(Boolean);
  if (caps.length < 2) return null;
  let n = 1;
  if (parts[1]) {
    const m = String(parts[1]).match(/\d+/);
    if (m) n = Math.max(1, Number(m[0]));
  }
  return {a:caps[0], b:caps[1], min:n};
}

function staffHasCapability_(staff, label) {
  const x = String(label || '').trim();
  if (x === '運転可') return !!staff.canDrive;
  if (x === '添乗可') return !!staff.canRide;
  if (x === '送迎可') return !!staff.canTransport;
  if (x === '直接支援配置可' || x === '直接支援配置算入') return !!staff.directSupport;
  if (x === '児童指導員配置可') return !!staff.legalEligible;
  if (x === '専門支援配置可') return !!staff.specialistEligible;
  if (x === '児発管配置可') return !!staff.jihatsukanEligible;
  if (x === '男性' || x === '女性') return String(staff.sex || '').trim() === x;
  return String(staff.role || '').trim() === x || String(staff.addOnCategory || '').trim() === x;
}

function ensureRoleCombinationRules_(ctx, state) {
  const rules = (ctx.rules && ctx.rules.facilityRules || []).filter(function (r) {
    return isRuleEnabled_(r) && String(r['ルール名'] || '').trim() === '役割組合せ';
  });
  rules.forEach(function (r) {
    const parsed = parseRoleCombinationRule_(r);
    if (!parsed) return;
    const target = normalizeUnit_(r['対象']);
    const units = target === '第1' || target === '第2' ? [target] : (ctx.unitCount >= 2 ? ['第1','第2'] : ['第1']);
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (ctx.days[d].state === '休日') continue;
      units.forEach(function (unit) {
        ensureUnitCapabilityCount_(ctx, state, d, unit, parsed.min, function (st) { return staffHasCapability_(st, parsed.a); }, parsed.a);
        ensureUnitCapabilityCount_(ctx, state, d, unit, parsed.min, function (st) { return staffHasCapability_(st, parsed.b); }, parsed.b);
      });
    }
  });
}

function ensureUnitCapabilityCount_(ctx, state, dayIndex, unit, needed, predicate, label) {
  function countNow() {
    let n = 0;
    ctx.staff.forEach(function (st) {
      if (!st.active || state.work[st.index][dayIndex] !== 1) return;
      if (getAssignedUnit_(ctx, state, st.index, dayIndex) !== unit) return;
      if (predicate(st)) n++;
    });
    return n;
  }
  let guard = 0;
  while (countNow() < needed && guard++ < 30) {
    const candidates = ctx.staff.filter(function (st) {
      if (!st.active || !predicate(st)) return false;
      if (!isUnitAllowedForStaff_(ctx, st, dayIndex, unit)) return false;
      if (hardUnavailableReason_(ctx, st, dayIndex)) return false;
      if (state.immutable[st.index][dayIndex] && state.work[st.index][dayIndex] !== 1) return false;
      if (state.work[st.index][dayIndex] !== 1 && wouldExceedMaxConsecutive_(state.work[st.index], dayIndex, getStaffMaxConsecutive_(ctx, st, dayIndex))) return false;
      return !(state.work[st.index][dayIndex] === 1 && getAssignedUnit_(ctx, state, st.index, dayIndex) === unit);
    });
    candidates.sort(function (a,b) {
      return candidateDayScore_(ctx,b,dayIndex,state) - candidateDayScore_(ctx,a,dayIndex,state);
    });
    if (!candidates.length) break;
    const st = candidates[0];
    state.work[st.index][dayIndex] = 1;
    setAssignedUnit_(ctx, state, st.index, dayIndex, unit);
  }
  if (countNow() < needed) {
    state.warnings.push((dayIndex + 1) + '日 ' + unit + '：役割「' + label + '」を' + needed + '名確保できません。');
  }
}

function patternExists_(ctx, patternName) {
  const rows = ctx.patternValues || [];
  if (!rows.length) return false;
  const headers = (rows[0] || []).map(String);
  const h = buildHeaderMap_(headers);
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (String(row[0] || '').trim() !== '○') continue;
    const name = String(valueAt_(row, h, '勤務パターン名') || valueAt_(row, h, 'パターン名') || row[1] || '').trim();
    if (name === patternName) return true;
  }
  return false;
}

function validateAdvancedRuleDefinitions_(ctx) {
  const errors = [];
  const warnings = [];
  (ctx.rules && ctx.rules.personRules || []).forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    const ref = String(r['対象職員'] || '').trim();
    if (!findStaffByRef_(ctx, ref)) errors.push('個人ルール ' + r._row + '行目：対象職員「' + ref + '」が見つかりません。');
    const p = numberOrZero_(r['優先度']);
    if (p < 0 || p > 100) warnings.push('個人ルール ' + r._row + '行目：優先度は0～100推奨です。');
    const type = String(r['ルール種別'] || '').trim();
    if ((type === '最大連勤' || type === '最低連休') && numberOrZero_(r['値1']) < 1) {
      errors.push('個人ルール ' + r._row + '行目：' + type + 'の値1は1以上にしてください。');
    }
    if (type === '勤務パターン指定') {
      const name = String(r['値1'] || '').trim();
      if (!name) {
        errors.push('個人ルール ' + r._row + '行目：勤務パターン指定の値1が空欄です。');
      } else if (!patternExists_(ctx, name)) {
        errors.push('個人ルール ' + r._row + '行目：勤務パターン「' + name + '」が見つかりません。');
      }
    }
  });
  (ctx.rules && ctx.rules.pairRules || []).forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    const a = String(r['職員A'] || '').trim();
    const b = String(r['職員B'] || '').trim();
    if (!findStaffByRef_(ctx, a)) errors.push('ペアルール ' + r._row + '行目：職員A「' + a + '」が見つかりません。');
    if (!findStaffByRef_(ctx, b)) errors.push('ペアルール ' + r._row + '行目：職員B「' + b + '」が見つかりません。');
    if (a && b && a === b) errors.push('ペアルール ' + r._row + '行目：同じ職員をA/Bに指定できません。');
  });
  (ctx.rules && ctx.rules.facilityRules || []).forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    if (String(r['ルール名'] || '').trim() === '役割組合せ' && !parseRoleCombinationRule_(r)) {
      warnings.push('施設ルール ' + r._row + '行目：役割組合せは「能力A+能力B:1」の形式で入力してください。');
    }
  });
  (ctx.rules && ctx.rules.exceptionRules || []).forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    if (String(r['ルール種別'] || '').trim() !== '一時的優先度変更') return;
    const target = String(r['対象'] || '').trim();
    const exists = (ctx.rules.facilityRules || []).some(function (fr) {
      return isRuleEnabled_(fr) && (String(fr['ルール名'] || '').trim() === target || String(fr['ルールID'] || '').trim() === target);
    });
    if (!exists) errors.push('月間・例外ルール ' + r._row + '行目：優先度変更の対象「' + target + '」が施設ルールに見つかりません。');
    const n = Number(r['値1']);
    if (!Number.isFinite(n) || n < 0 || n > 100) errors.push('月間・例外ルール ' + r._row + '行目：一時的優先度変更の値1は0～100にしてください。');
  });
  return {errors:unique_(errors), warnings:unique_(warnings)};
}

function validateAdvancedPlan_(ctx, plan) {
  const errors = [];
  const warnings = [];

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    for (let d = 0; d < ctx.daysInMonth; d++) {
      const works = plan.work[staff.index][d] === 1;
      const unit = getAssignedUnitFromMatrices_(staff, plan.unit, staff.index, d);
      getPersonRules_(ctx, staff, null, d).forEach(function (r) {
        const type = String(r['ルール種別'] || '').trim();
        const strength = String(r['強度'] || '').trim();
        let violated = false;
        let text = '';
        if (type === '出勤固定' && !works) { violated = true; text = '出勤固定'; }
        if (type === '休み固定' && works) { violated = true; text = '休み固定'; }
        if (type === 'Unit禁止' && works && normalizeUnit_(r['値1']) === unit) { violated = true; text = 'Unit禁止'; }
        if (type === 'Unit優先' && works && normalizeUnit_(r['値1']) && normalizeUnit_(r['値1']) !== unit) { violated = true; text = 'Unit優先'; }
        if (violated) {
          const msg = (d + 1) + '日 ' + staff.name + '：個人ルール「' + text + '」未達';
          (strength === '絶対' ? errors : warnings).push(msg);
        }
      });
    }

    const minRestRules = getPersonRules_(ctx, staff, '最低連休', null);
    minRestRules.forEach(function (r) {
      const n = numberOrZero_(r['値1']);
      if (n > 0 && !hasConsecutiveOff_(plan.work[staff.index], ctx.daysInMonth, n)) {
        const msg = staff.name + '：最低連休' + n + '日を確保できていません。';
        (String(r['強度'] || '') === '絶対' ? errors : warnings).push(msg);
      }
    });
  });

  const pairRows = ctx.rules && ctx.rules.pairRules || [];
  pairRows.forEach(function (r) {
    if (!isRuleEnabled_(r)) return;
    const a = findStaffByRef_(ctx, r['職員A']);
    const b = findStaffByRef_(ctx, r['職員B']);
    if (!a || !b) return;
    const type = String(r['ルール'] || '').trim();
    let satisfied = 0;
    let applicable = 0;
    const badDays = [];
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (!ruleDateMatches_(ctx, r, d)) continue;
      applicable++;
      const aw = plan.work[a.index][d] === 1;
      const bw = plan.work[b.index][d] === 1;
      const au = aw ? getAssignedUnitFromMatrices_(a, plan.unit, a.index, d) : '';
      const bu = bw ? getAssignedUnitFromMatrices_(b, plan.unit, b.index, d) : '';
      let ok = true;
      if (type === '合わせる') ok = aw === bw;
      if (type === 'どちらか一方') ok = aw !== bw;
      if (type === '合わせない') ok = !(aw && bw);
      if (type === '同時休み優先') ok = !(aw !== bw);
      if (type === '同時休み禁止') ok = aw || bw;
      if (type === '同一Unit優先' && aw && bw) ok = au === bu;
      if (type === '別Unit優先' && aw && bw) ok = au !== bu;
      if (ok) satisfied++; else badDays.push(d + 1);
    }
    const min = numberOrZero_(r['月最低回数']);
    const violates = min > 0 ? satisfied < min : badDays.length > 0;
    if (violates) {
      const msg = 'ペアルール「' + a.name + '×' + b.name + ' / ' + type + '」未達' +
        (min > 0 ? '（達成' + satisfied + '回 / 最低' + min + '回）' : '（' + badDays.slice(0,10).join('・') + '日）');
      (String(r['強度'] || '') === '絶対' ? errors : warnings).push(msg);
    }
  });

  (ctx.rules && ctx.rules.facilityRules || []).filter(function (r) {
    return isRuleEnabled_(r) && String(r['ルール名'] || '').trim() === '役割組合せ';
  }).forEach(function (r) {
    const parsed = parseRoleCombinationRule_(r);
    if (!parsed) return;
    const target = normalizeUnit_(r['対象']);
    const units = target === '第1' || target === '第2' ? [target] : (ctx.unitCount >= 2 ? ['第1','第2'] : ['第1']);
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (ctx.days[d].state === '休日') continue;
      units.forEach(function (unit) {
        let a = 0, b = 0;
        ctx.staff.forEach(function (st) {
          if (!st.active || plan.work[st.index][d] !== 1) return;
          if (getAssignedUnitFromMatrices_(st, plan.unit, st.index, d) !== unit) return;
          if (staffHasCapability_(st, parsed.a)) a++;
          if (staffHasCapability_(st, parsed.b)) b++;
        });
        if (a < parsed.min || b < parsed.min) {
          const msg = (d + 1) + '日 ' + unit + '：役割組合せ「' + parsed.a + '+' + parsed.b + '」が不足しています。';
          (String(r['強度'] || '') === '絶対' ? errors : warnings).push(msg);
        }
      });
    }
  });

  if (facilityRuleEnabled_(ctx.rules, '土日祝公平性')) {
    const counts = ctx.staff.filter(function (st) { return st.active && st.employment === '社員'; }).map(function (st) {
      let n = 0;
      for (let d = 0; d < ctx.daysInMonth; d++) if (isWeekendLikeDay_(ctx.days[d]) && plan.work[st.index][d] === 1) n++;
      return {name:st.name, count:n};
    });
    if (counts.length >= 2) {
      const nums = counts.map(function (x) { return x.count; });
      const max = Math.max.apply(null, nums), min = Math.min.apply(null, nums);
      if (max - min > 1) warnings.push('土日祝勤務の偏りがあります（最大' + max + '回 / 最小' + min + '回）。');
    }
  }

  for (let d = 0; d < ctx.daysInMonth; d++) {
    if (ctx.days[d].state === '休日') continue;
    ['第1','第2'].slice(0,ctx.unitCount).forEach(function (unit) {
      const need = requiredDirectSupportCount_(ctx, ctx.days[d], unit);
      const have = countLegalStaffFromPlan_(ctx, plan, d, unit);
      if (have < need) {
        const r = facilityRule_(ctx.rules, '最低現場余裕人数');
        const msg = (d + 1) + '日 ' + unit + '：現場余裕を含む必要人数' + need + '名に対し' + have + '名です。';
        (r && String(r['強度'] || '') === '絶対' ? errors : warnings).push(msg);
      }
    });
  }

  return {errors:unique_(errors), warnings:unique_(warnings)};
}

function hasConsecutiveOff_(workRow, daysInMonth, needed) {
  let run = 0;
  for (let d = 0; d < daysInMonth; d++) {
    if (workRow[d] === 1) run = 0; else run++;
    if (run >= needed) return true;
  }
  return false;
}

function ensureConfirmedHistorySheet_(ss) {
  let sh = ss.getSheetByName(SHIFT_APP.SHEETS.CONFIRMED_HISTORY);
  if (!sh) {
    sh = ss.insertSheet(SHIFT_APP.SHEETS.CONFIRMED_HISTORY);
    sh.getRange('A1:F1').setValues([['確定日時','対象月','スナップショット','警告数','警告内容','実行者']]);
  }
  sh.hideSheet();
  return sh;
}

function ensureActualLogSheet_(ss) {
  let sh = ss.getSheetByName(SHIFT_APP.SHEETS.ACTUAL_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHIFT_APP.SHEETS.ACTUAL_LOG);
    sh.getRange('A1:I1').setValues([['記録日時','対象月','職員ID','氏名','休みにした日','出勤にした日','Unit','メモ','実行者']]);
  }
  sh.hideSheet();
  return sh;
}

function appendActualChangeLog_(ss, ctx, staff, offDate, onDate, unit, memo) {
  const sh = ensureActualLogSheet_(ss);
  sh.appendRow([new Date(), ctx.targetMonth, staff.id, staff.name, offDate, onDate, unit || '', memo || '', Session.getActiveUser().getEmail() || '']);
}

function assertDateInTargetMonth_(ctx, date, label) {
  if (date.getFullYear() !== ctx.targetMonth.getFullYear() || date.getMonth() !== ctx.targetMonth.getMonth()) {
    throw new Error(label + 'が対象月外です：' + Utilities.formatDate(date, ctx.ss.getSpreadsheetTimeZone(), 'yyyy/M/d'));
  }
}

function addGenericLock_(ctx, date, staffId, type, note) {
  const sh = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.LOCK);
  const values = sh.getRange('A5:H80').getValues();
  let row = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === '') { row = i + 5; break; }
  }
  if (row < 0) throw new Error('「配置ロック」に空き行がありません。');
  sh.getRange(row,1,1,3).setValues([['○',date,staffId]]);
  sh.getRange(row,5,1,3).setValues([[type,'',note || '']]);
}


function clearExpiredMonthlyExceptionRules_(sheet, sourceMonth) {
  const range = sheet.getRange('A5:L220');
  const values = range.getValues();
  const cutoff = new Date(sourceMonth.getFullYear(), sourceMonth.getMonth() + 1, 1);
  for (let r = 0; r < values.length; r++) {
    const start = normalizeDate_(values[r][3]);
    const end = normalizeDate_(values[r][4]);
    const reference = end || start;
    if (!reference || reference.getTime() < cutoff.getTime()) {
      sheet.getRange(r + 5, 1, 1, 12).clearContent();
    }
  }
}

function setLabelValue_(sheet, label, value) {
  const range = sheet.getRange('A1:D100');
  const vals = range.getValues();
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length - 1; c++) {
      if (String(vals[r][c] || '').trim() === label) {
        sheet.getRange(r + 1, c + 2).setValue(value);
        return true;
      }
    }
  }
  throw new Error('設定項目「' + label + '」が見つかりません。');
}

/* =========================================================
 * 5. 初期設定検証
 * ========================================================= */

function validateInitialSetup_(ctx) {
  const errors = [];
  const warnings = [];

  const setupSheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.SETUP);
  const setup = setupSheet.getRange('A1:D80').getValues();

  if (!String(findLabelValue_(setup, '施設名') || '').trim()) {
    errors.push('施設名が未入力です。');
  }

  if (String(findLabelValue_(setup, 'サービス種別') || '').trim() !==
      '放課後等デイサービス') {
    errors.push('v1.1.2は「放課後等デイサービス」のみ対応しています。');
  }

  if (ctx.unitCount !== 1 && ctx.unitCount !== 2) {
    errors.push('単位数は1または2に設定してください。');
  }

  if (ctx.maxConsecutive < ctx.preferredConsecutive) {
    errors.push('最大連勤数が原則連勤数より小さくなっています。');
  }

  const activeStaff = ctx.staff.filter(function (x) { return x.active; });

  if (activeStaff.length === 0) {
    errors.push('在籍職員が1名も登録されていません。');
  }

  activeStaff.forEach(function (staff) {
    if (!staff.employment) {
      errors.push(staff.id + ' ' + staff.name + '：雇用区分が未入力です。');
    }

    if (staff.employment === '社員' && !staff.workStyle) {
      errors.push(staff.id + ' ' + staff.name + '：働き方が未設定です。');
    }

    if (staff.employment === '社員' && staff.requiredDays <= 0) {
      errors.push(
        staff.id + ' ' + staff.name +
        '：必要勤務日数を計算できていません。働き方設定を確認してください。'
      );
    }

    if (staff.restrictionMemo) {
      warnings.push(
        staff.id + ' ' + staff.name +
        '：「勤務制限・備考」の自由記述は自動判定しません。' +
        '必須条件は「固定勤務曜日」「固定休日」「配置ロック」で構造化してください。'
      );
    }
  });

  for (let u = 1; u <= ctx.unitCount; u++) {
    const unit = '第' + u;
    const config = ctx.placement[unit];

    if (!config || !config.jihatsukan ||
        (!config.jihatsukan.main && !config.jihatsukan.sub)) {
      errors.push(unit + '単位の児発管担当者が未設定です。');
      continue;
    }

    const managerRefs = [
      config.jihatsukan.main,
      config.jihatsukan.sub
    ].filter(Boolean);

    managerRefs.forEach(function (ref) {
      const staff = findStaffByRef_(ctx, ref);
      if (!staff) {
        errors.push(unit + '単位の児発管「' + ref + '」が職員登録に存在しません。');
      } else if (!staff.jihatsukanEligible) {
        errors.push(
          unit + '単位の児発管「' + staff.name +
          '」は「児発管配置可」が○ではありません。'
        );
      }
    });
  }

  // 同じ主担当が第1・第2の両方に指定されている場合は、同時運営時に矛盾し得る。
  if (ctx.unitCount === 2) {
    const a = ctx.placement['第1'] &&
      ctx.placement['第1'].jihatsukan &&
      ctx.placement['第1'].jihatsukan.main;
    const b = ctx.placement['第2'] &&
      ctx.placement['第2'].jihatsukan &&
      ctx.placement['第2'].jihatsukan.main;

    if (a && b && a === b) {
      warnings.push(
        '第1・第2の児発管主担当が同一です。両単位を同時運営する日は副担当等の確認が必要です。'
      );
    }
  }

  const enabledStyles = ctx.workStyleValues.slice(1).filter(function (row) {
    return String(row[0] || '').trim() === '○' &&
           String(row[1] || '').trim() !== '';
  });

  if (enabledStyles.length === 0) {
    errors.push('使用する働き方が1件も登録されていません。');
  }

  const enabledPatterns = ctx.patternValues.slice(1).filter(function (row) {
    return String(row[0] || '').trim() === '○' &&
           String(row[1] || '').trim() !== '';
  });

  if (enabledPatterns.length < 2) {
    errors.push(
      '勤務パターンが不足しています。少なくとも平日と学校休日の通常勤務を登録してください。'
    );
  }

  const adminCheck = validateTokyoAdminCompliance_(ctx);
  Array.prototype.push.apply(errors, adminCheck.errors);
  Array.prototype.push.apply(warnings, adminCheck.warnings);

  if (ctx.transportUse === '使用する') {
    if (ctx.transportVehicleCount < 1) {
      errors.push('送迎を使用する場合、通常運用車両数を1台以上にしてください。');
    }

    const driverCount = activeStaff.filter(function (x) {
      return x.canDrive;
    }).length;

    const employeeTransportCount = activeStaff.filter(function (x) {
      return x.employment === '社員' && x.canTransport;
    }).length;

    if (driverCount < ctx.transportVehicleCount) {
      warnings.push(
        '運転可能者が通常運用車両数より少ないため、毎日の送迎成立が難しい可能性があります。'
      );
    }

    if (ctx.employeeRideRule === '必須' &&
        employeeTransportCount < ctx.transportVehicleCount) {
      warnings.push(
        '社員同乗が「必須」ですが、送迎可能な社員数が通常運用車両数より少ないです。'
      );
    }

    if (ctx.employeeRideRule === '施設ルールによる' || !ctx.employeeRideRule) {
      errors.push(
        '送迎の「各車両の社員同乗」が未確定です。「必須」または「不要」を選択してください。'
      );
    }
  }

  // ロックの入力矛盾
  const lockCheck = validateLocks_(ctx);
  Array.prototype.push.apply(errors, lockCheck.errors);
  Array.prototype.push.apply(warnings, lockCheck.warnings);

  const advanced = validateAdvancedRuleDefinitions_(ctx);
  Array.prototype.push.apply(errors, advanced.errors);
  Array.prototype.push.apply(warnings, advanced.warnings);

  return {
    errors: unique_(errors),
    warnings: unique_(warnings)
  };
}


/* =========================================================
 * 6. 初期状態・スコープ
 * ========================================================= */

function createInitialState_(ctx, existing, options) {
  const work = matrix_(SHIFT_APP.MAX_STAFF, SHIFT_APP.MAX_DAYS, 0);
  const unit = matrix_(SHIFT_APP.MAX_STAFF, SHIFT_APP.MAX_DAYS, '');
  const immutable = matrix_(SHIFT_APP.MAX_STAFF, SHIFT_APP.MAX_DAYS, false);
  const lockType = matrix_(SHIFT_APP.MAX_STAFF, SHIFT_APP.MAX_DAYS, '');
  const role = matrix_(SHIFT_APP.MAX_STAFF, SHIFT_APP.MAX_DAYS, '');

  const scope = options.scope || { type: 'all' };

  if (scope.type !== 'all') {
    for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
      for (let d = 0; d < SHIFT_APP.MAX_DAYS; d++) {
        if (d >= ctx.daysInMonth) {
          immutable[s][d] = true;
          continue;
        }

        const mutable = isCellInScope_(
          ctx,
          existing,
          s,
          d,
          scope
        );

        if (!mutable) {
          work[s][d] = existing.work[s][d] ? 1 : 0;
          unit[s][d] = existing.unit[s][d] || '';
          immutable[s][d] = true;
          lockType[s][d] = '部分再生成の範囲外';
        }
      }
    }
  }

  // 再最適化・部分再生成では「自動調整禁止」のセルを現在状態のまま固定する。
  if (options.mode !== 'generate') {
    for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
      const staff = ctx.staff[s];
      if (!staff || !staff.active) continue;
      for (let d = 0; d < ctx.daysInMonth; d++) {
        if (!isAutoAdjustBlocked_(ctx, staff, d)) continue;
        work[s][d] = existing.work[s][d] ? 1 : 0;
        unit[s][d] = existing.unit[s][d] || '';
        immutable[s][d] = true;
        lockType[s][d] = '自動調整禁止';
      }
    }
  }

  return {
    work: work,
    unit: unit,
    immutable: immutable,
    lockType: lockType,
    role: role,
    warnings: [],
    errors: []
  };
}


function isCellInScope_(ctx, existing, staffIndex, dayIndex, scope) {
  if (scope.type === 'all') {
    return true;
  }

  if (scope.type === 'day') {
    return dayIndex === scope.day - 1;
  }

  if (scope.type === 'staff') {
    return ctx.staff[staffIndex].id === scope.staffId;
  }

  if (scope.type === 'unit') {
    const staff = ctx.staff[staffIndex];
    const assigned = getAssignedUnitFromMatrices_(
      staff,
      existing.unit,
      staffIndex,
      dayIndex
    );

    // 現在その単位にいる人、または元所属がその単位の人は再生成対象。
    if (assigned === scope.unit) {
      return true;
    }

    if (!existing.work[staffIndex][dayIndex] &&
        staff.unit === scope.unit) {
      return true;
    }

    return false;
  }

  return true;
}


/* =========================================================
 * 7. ハード条件
 * ========================================================= */

function enforceHardAvailability_(ctx, state) {
  for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
    const staff = ctx.staff[s];

    for (let d = 0; d < SHIFT_APP.MAX_DAYS; d++) {
      if (state.immutable[s][d]) {
        continue;
      }

      if (d >= ctx.daysInMonth || !staff.active) {
        state.work[s][d] = 0;
        state.unit[s][d] = '';
        continue;
      }

      const reason = hardUnavailableReason_(ctx, staff, d);

      if (reason) {
        state.work[s][d] = 0;
        state.unit[s][d] = '';
        state.lockType[s][d] = reason;
      }
    }
  }
}


function hardUnavailableReason_(ctx, staff, dayIndex) {
  if (!staff.active) {
    return '未在籍';
  }

  if (dayIndex >= ctx.daysInMonth) {
    return '対象月外';
  }

  const day = ctx.days[dayIndex];
  const pref = staff.preferences[dayIndex];

  if (day.state === '休日') {
    return '施設休業日';
  }

  if (pref === '不可') {
    return '勤務不可';
  }

  if (pref === '有休') {
    return '有休';
  }

  const advancedOff = advancedHardOffReason_(ctx, staff, dayIndex);
  if (advancedOff) {
    return advancedOff;
  }

  const weekday = day.weekday;

  const fixedOff = parseWeekdays_(staff.fixedDaysOff);
  if (fixedOff.length > 0 && fixedOff.indexOf(weekday) >= 0) {
    return '固定休日';
  }

  const fixedWork = parseWeekdays_(staff.fixedWorkWeekdays);
  if (fixedWork.length > 0 && fixedWork.indexOf(weekday) < 0) {
    return '固定勤務曜日外';
  }

  // パート・ドライバーの「出勤希望」は勤務可能日の意味。
  if (staff.employment !== '社員' && pref !== '出勤希望') {
    return '出勤希望なし';
  }

  return '';
}


/* =========================================================
 * 8. 配置ロック
 * ========================================================= */

function parseLocks_(values, targetMonth, daysInMonth) {
  const headers = (values[0] || []).map(String);
  const header = buildHeaderMap_(headers);
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    if (String(valueAt_(row, header, '使用') || '').trim() !== '○') {
      continue;
    }

    const dateValue = valueAt_(row, header, '日付');
    const staffId = String(valueAt_(row, header, '職員ID') || '').trim();
    const type = String(valueAt_(row, header, 'ロック種別') || '').trim();
    const note = String(valueAt_(row, header, '理由・備考') || '').trim();

    const date = normalizeDate_(dateValue);
    if (!date || isNaN(date.getTime())) {
      result.push({
        row: i + 4,
        invalid: true,
        reason: '日付が不正です。'
      });
      continue;
    }

    if (date.getFullYear() !== targetMonth.getFullYear() ||
        date.getMonth() !== targetMonth.getMonth()) {
      // 対象月以外のロックは保持するが、今月の生成には使用しない。
      continue;
    }

    const day = date.getDate();
    if (day < 1 || day > daysInMonth) {
      continue;
    }

    result.push({
      row: i + 4,
      date: date,
      day: day,
      staffId: staffId,
      type: type,
      note: note
    });
  }

  return result;
}


function validateLocks_(ctx) {
  const errors = [];
  const warnings = [];
  const grouped = {};

  ctx.locks.forEach(function (lock) {
    if (lock.invalid) {
      errors.push('配置ロック ' + lock.row + '行目：' + lock.reason);
      return;
    }

    const staff = ctx.staff.find(function (x) {
      return x.id === lock.staffId && x.active;
    });

    if (!staff) {
      errors.push(
        '配置ロック ' + lock.row + '行目：在籍職員ID「' +
        lock.staffId + '」が見つかりません。'
      );
      return;
    }

    if (['出勤固定', '休み固定', '第1固定', '第2固定']
        .indexOf(lock.type) < 0) {
      errors.push(
        '配置ロック ' + lock.row + '行目：ロック種別が不正です。'
      );
      return;
    }

    if (lock.type === '第2固定' && ctx.unitCount < 2) {
      errors.push(
        '配置ロック ' + lock.row + '行目：1単位施設では第2固定を使用できません。'
      );
    }

    const key = lock.staffId + '_' + lock.day;
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(lock.type);
  });

  Object.keys(grouped).forEach(function (key) {
    const types = unique_(grouped[key]);

    if (types.indexOf('休み固定') >= 0 &&
        types.some(function (x) {
          return x === '出勤固定' || x === '第1固定' || x === '第2固定';
        })) {
      errors.push(
        '配置ロックに同一職員・同一日の「休み」と「出勤」の矛盾があります：' + key
      );
    }

    if (types.indexOf('第1固定') >= 0 &&
        types.indexOf('第2固定') >= 0) {
      errors.push(
        '配置ロックに同一職員・同一日の第1固定と第2固定の矛盾があります：' + key
      );
    }
  });

  return { errors: errors, warnings: warnings };
}


function applyPlacementLocks_(ctx, state) {
  const errors = [];
  const warnings = [];

  ctx.locks.forEach(function (lock) {
    if (lock.invalid) {
      return;
    }

    const staffIndex = ctx.staff.findIndex(function (x) {
      return x.id === lock.staffId && x.active;
    });

    if (staffIndex < 0) {
      return;
    }

    const dayIndex = lock.day - 1;
    const staff = ctx.staff[staffIndex];
    const unavailable = hardUnavailableReason_(ctx, staff, dayIndex);

    if (lock.type === '休み固定') {
      if (!state.immutable[staffIndex][dayIndex]) {
        state.work[staffIndex][dayIndex] = 0;
        state.unit[staffIndex][dayIndex] = '';
      }
      state.immutable[staffIndex][dayIndex] = true;
      state.lockType[staffIndex][dayIndex] = '休み固定';
      return;
    }

    if (unavailable) {
      errors.push(
        lock.day + '日 ' + staff.name + '：' +
        lock.type + 'ですが「' + unavailable + '」と矛盾します。'
      );
      return;
    }

    if (state.immutable[staffIndex][dayIndex] &&
        state.work[staffIndex][dayIndex] !== 1) {
      errors.push(
        lock.day + '日 ' + staff.name +
        '：部分再生成の固定状態と配置ロックが矛盾します。'
      );
      return;
    }

    state.work[staffIndex][dayIndex] = 1;
    state.immutable[staffIndex][dayIndex] = true;
    state.lockType[staffIndex][dayIndex] = lock.type;

    if (lock.type === '第1固定') {
      state.unit[staffIndex][dayIndex] = '第1';
    } else if (lock.type === '第2固定') {
      state.unit[staffIndex][dayIndex] = '第2';
    }
  });

  return {
    errors: unique_(errors),
    warnings: unique_(warnings)
  };
}


/* =========================================================
 * 9. 初期勤務日の選択
 * ========================================================= */

function seedEmployeeWorkDays_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') {
      return;
    }

    const target = getWorkTargetDays_(staff);

    let current = countWorkDays_(state.work[staff.index], ctx.daysInMonth);

    if (current >= target) {
      return;
    }

    const candidates = [];

    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (state.work[staff.index][d] === 1) continue;
      if (state.immutable[staff.index][d]) continue;
      if (hardUnavailableReason_(ctx, staff, d)) continue;

      candidates.push({
        dayIndex: d,
        score: candidateDayScore_(ctx, staff, d, state)
      });
    }

    candidates.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.dayIndex - b.dayIndex;
    });

    for (let i = 0; i < candidates.length && current < target; i++) {
      const d = candidates[i].dayIndex;

      if (wouldExceedMaxConsecutive_(
        state.work[staff.index],
        d,
        getStaffMaxConsecutive_(ctx, staff, d)
      )) {
        continue;
      }

      state.work[staff.index][d] = 1;
      current++;
    }

    if (current < target) {
      state.warnings.push(
        staff.name + '：最大連勤・不可日等を守ると、必要勤務日数 ' +
        target + '日に対して ' + current + '日しか確保できません。'
      );
    }
  });
}


function getWorkTargetDays_(staff) {
  if (!staff.active || staff.employment !== '社員') {
    return 0;
  }

  return Math.max(
    0,
    Math.round(numberOrZero_(staff.requiredDays)) -
    numberOrZero_(staff.paidLeaveDays)
  );
}


function candidateDayScore_(ctx, staff, dayIndex, state) {
  let score = numberOrZero_(staff.scores[dayIndex]);
  const pref = staff.preferences[dayIndex];

  if (pref === '希望休') score -= 100000;
  if (pref === '出勤希望') score += 10000;

  score += personRuleDayScore_(ctx, staff, dayIndex);
  score += exceptionRuleDayScore_(ctx, staff, dayIndex);
  score += pairRuleDayScore_(ctx, staff, dayIndex, state);

  const day = ctx.days[dayIndex];
  if (state && isWeekendLikeDay_(day) && facilityRuleEnabled_(ctx.rules, '土日祝公平性')) {
    const weight = facilityRulePriorityForDay_(ctx, '土日祝公平性', dayIndex, 40);
    score -= countWeekendLikeWork_(ctx, state, staff.index) * weight * 20;
  }

  if (state && facilityRuleEnabled_(ctx.rules, '休み分散')) {
    const weight = facilityRulePriorityForDay_(ctx, '休み分散', dayIndex, 35);
    const left = dayIndex > 0 ? state.work[staff.index][dayIndex - 1] : 0;
    const right = dayIndex + 1 < ctx.daysInMonth ? state.work[staff.index][dayIndex + 1] : 0;
    if (left && right) score -= weight * 30;
    else if (left || right) score -= weight * 8;
  }

  if (facilityRuleEnabled_(ctx.rules, '人件費最適化') && staff.costPerHour > 0) {
    const weight = facilityRulePriorityForDay_(ctx, '人件費最適化', dayIndex, 20);
    score -= Math.min(5000, staff.costPerHour) * weight / 100;
  }

  return score;
}


/* =========================================================
 * 10. 児発管
 * ========================================================= */

function ensureJihatsukan_(ctx, state) {
  for (let d = 0; d < ctx.daysInMonth; d++) {
    for (let u = 1; u <= ctx.unitCount; u++) {
      const unit = '第' + u;
      const day = ctx.days[d];
      const usage = u === 1 ? day.usage1 : day.usage2;
      const legalNeed = u === 1 ? day.legal1 : day.legal2;

      if (day.state === '休日' || (usage <= 0 && legalNeed <= 0)) {
        continue;
      }

      const config = ctx.placement[unit];
      if (!config || !config.jihatsukan) {
        state.errors.push(d + 1 + '日 ' + unit + '：児発管設定がありません。');
        continue;
      }

      const refs = [
        config.jihatsukan.main,
        config.jihatsukan.sub
      ].filter(Boolean);

      const candidates = [];

      for (let r = 0; r < refs.length; r++) {
        const staff = findStaffByRef_(ctx, refs[r]);

        if (!staff || !staff.active || !staff.jihatsukanEligible) {
          continue;
        }

        if (hardUnavailableReason_(ctx, staff, d)) {
          continue;
        }

        // 部分再生成の範囲外は変更しない。
        if (state.immutable[staff.index][d]) {
          if (state.work[staff.index][d] !== 1) {
            continue;
          }

          const fixedUnit = getAssignedUnit_(ctx, state, staff.index, d);
          if (fixedUnit && fixedUnit !== unit) {
            continue;
          }
        }

        // 同日に別単位の児発管として既に固定されている場合は使わない。
        const currentRole = state.role[staff.index][d];
        if (currentRole.indexOf('児発管:') === 0 &&
            currentRole !== '児発管:' + unit) {
          continue;
        }

        const alreadyWorking = state.work[staff.index][d] === 1;
        let swapOutDay = -1;

        if (!alreadyWorking) {
          if (wouldExceedMaxConsecutive_(
            state.work[staff.index],
            d,
            getStaffMaxConsecutive_(ctx, staff, d)
          )) {
            continue;
          }

          // 社員は、児発管確保のためだけに契約日数を増やさない。
          // 既に必要勤務日数へ達している場合は、別の可変勤務日と入れ替える。
          if (staff.employment === '社員') {
            const target = getWorkTargetDays_(staff);
            const current = countWorkDays_(
              state.work[staff.index],
              ctx.daysInMonth
            );

            if (current >= target) {
              swapOutDay = findJihatsukanSwapOutDay_(
                ctx,
                state,
                staff,
                d
              );

              if (swapOutDay < 0) {
                continue;
              }
            }
          }
        }

        let priority = alreadyWorking ? 1000000 : 100000;

        // 主担当を基本優先。ただし「既に勤務している副担当」は、
        // 休みの主担当を追加出勤させるより優先する。
        priority -= r * 1000;

        // 希望休への勤務は最後の手段にする。
        if (staff.preferences[d] === '希望休') {
          priority -= 500000;
        }

        candidates.push({
          staff: staff,
          priority: priority,
          alreadyWorking: alreadyWorking,
          swapOutDay: swapOutDay
        });
      }

      candidates.sort(function (a, b) {
        if (b.priority !== a.priority) return b.priority - a.priority;
        return a.staff.index - b.staff.index;
      });

      let placed = false;

      for (let c = 0; c < candidates.length; c++) {
        const candidate = candidates[c];
        const staff = candidate.staff;

        if (!candidate.alreadyWorking && candidate.swapOutDay >= 0) {
          const out = candidate.swapOutDay;
          state.work[staff.index][out] = 0;
          state.unit[staff.index][out] = '';
          state.role[staff.index][out] = '';
        }

        state.work[staff.index][d] = 1;
        setAssignedUnit_(ctx, state, staff.index, d, unit);
        state.role[staff.index][d] = '児発管:' + unit;
        placed = true;
        break;
      }

      if (!placed) {
        state.errors.push(
          (d + 1) + '日 ' + unit +
          '：契約日数・勤務条件を守ったまま児発管を配置できません。'
        );
      }
    }
  }
}


/**
 * 児発管として新たに勤務させる日の代わりに外せる勤務日を探す。
 * 既に役割が付いている日、固定日、希望出勤日は極力残す。
 */
function findJihatsukanSwapOutDay_(ctx, state, staff, incomingDayIndex) {
  const removable = [];

  for (let d = 0; d < ctx.daysInMonth; d++) {
    if (d === incomingDayIndex) continue;
    if (state.work[staff.index][d] !== 1) continue;
    if (state.immutable[staff.index][d]) continue;
    if (state.role[staff.index][d]) continue;

    let score = candidateDayScore_(ctx, staff, d, state);

    // 出勤希望はできるだけ残す。
    if (staff.preferences[d] === '出勤希望') {
      score += 200000;
    }

    removable.push({
      dayIndex: d,
      score: score
    });
  }

  if (removable.length === 0) {
    return -1;
  }

  // スコアが低い日から外す。希望休が誤って勤務になっていれば最優先で外れる。
  removable.sort(function (a, b) {
    if (a.score !== b.score) return a.score - b.score;
    return b.dayIndex - a.dayIndex;
  });

  return removable[0].dayIndex;
}


/* =========================================================
 * 11. 単位割当
 * ========================================================= */

function assignUnassignedScheduledStaff_(ctx, state) {
  for (let d = 0; d < ctx.daysInMonth; d++) {
    for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
      const staff = ctx.staff[s];

      if (!staff.active || state.work[s][d] !== 1) {
        continue;
      }

      const assigned = getAssignedUnit_(ctx, state, s, d);

      if (assigned === '第1' || assigned === '第2') {
        continue;
      }

      // 共通所属等は、その日の法定不足が大きい単位へ先に置く。
      const deficit1 = Math.max(
        0,
        ctx.days[d].legal1 - countLegalStaff_(ctx, state, d, '第1')
      );

      const deficit2 = ctx.unitCount >= 2
        ? Math.max(
            0,
            ctx.days[d].legal2 - countLegalStaff_(ctx, state, d, '第2')
          )
        : -1;

      const fallback = ctx.unitCount === 1 || deficit1 >= deficit2 ? '第1' : '第2';
      const chosen = choosePreferredUnitForStaff_(ctx, state, staff, d, fallback);
      if (chosen) setAssignedUnit_(ctx, state, s, d, chosen);
    }
  }
}


/* =========================================================
 * 12. 法定配置
 * ========================================================= */

function ensureLegalStaffing_(ctx, state) {
  for (let d = 0; d < ctx.daysInMonth; d++) {
    const day = ctx.days[d];

    if (day.state === '休日') {
      continue;
    }

    for (let u = 1; u <= ctx.unitCount; u++) {
      const unit = '第' + u;
      const needed = requiredDirectSupportCount_(ctx, day, unit);

      if (needed <= 0) {
        continue;
      }

      let guard = 0;

      while (countLegalStaff_(ctx, state, d, unit) < needed &&
             guard < 50) {
        guard++;

        const candidate = findBestLegalCandidate_(
          ctx,
          state,
          d,
          unit
        );

        if (!candidate) {
          break;
        }

        if (candidate.moveFromOtherUnit) {
          setAssignedUnit_(
            ctx,
            state,
            candidate.staff.index,
            d,
            unit
          );
        } else {
          state.work[candidate.staff.index][d] = 1;
          setAssignedUnit_(
            ctx,
            state,
            candidate.staff.index,
            d,
            unit
          );
        }
      }

      const after = countLegalStaff_(ctx, state, d, unit);

      if (after < needed) {
        state.errors.push(
          (d + 1) + '日 ' + unit +
          '：法定必要 ' + needed + '名に対し、配置可能 ' +
          after + '名です。'
        );
      }
    }
  }
}


function countLegalStaff_(ctx, state, dayIndex, unit) {
  let count = 0;

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    if (state.work[staff.index][dayIndex] !== 1) return;
    if (!staff.directSupport || !staff.legalEligible) return;
    if (isWoodyTokyoMunicipality_(ctx) && !isTokyoCoreLegalStaff_(staff)) return;
    if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;

    const assigned = getAssignedUnit_(
      ctx,
      state,
      staff.index,
      dayIndex
    );

    if (assigned !== unit) return;

    // 同単位の児発管担当者は直接支援の法定人数から除外。
    if (state.role[staff.index][dayIndex] === '児発管:' + unit) {
      return;
    }

    count++;
  });

  return count;
}


function findBestLegalCandidate_(ctx, state, dayIndex, unit) {
  const candidates = [];

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    if (!staff.directSupport || !staff.legalEligible) return;
    if (isWoodyTokyoMunicipality_(ctx) && !isTokyoCoreLegalStaff_(staff)) return;
    if (!isUnitAllowedForStaff_(ctx, staff, dayIndex, unit)) return;
    if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;
    if (hardUnavailableReason_(ctx, staff, dayIndex)) return;

    const role = state.role[staff.index][dayIndex];

    if (role.indexOf('児発管:') === 0) {
      return;
    }

    const isWorking = state.work[staff.index][dayIndex] === 1;
    const assigned = isWorking
      ? getAssignedUnit_(ctx, state, staff.index, dayIndex)
      : '';

    if (isWorking && assigned === unit) {
      return;
    }

    // 部分再生成の範囲外は、勤務の追加だけでなく単位移動も禁止する。
    if (state.immutable[staff.index][dayIndex]) {
      return;
    }

    if (!isWorking &&
        wouldExceedMaxConsecutive_(
          state.work[staff.index],
          dayIndex,
          getStaffMaxConsecutive_(ctx, staff, dayIndex)
        )) {
      return;
    }

    let moveFromOtherUnit = false;

    if (isWorking && assigned && assigned !== unit) {
      // 他単位から動かしても、その単位が法定不足にならない場合だけ候補。
      const otherNeed = requiredDirectSupportCount_(ctx, ctx.days[dayIndex], assigned);

      const otherCount = countLegalStaff_(
        ctx,
        state,
        dayIndex,
        assigned
      );

      if (otherCount - 1 < otherNeed) {
        return;
      }

      // 配置ロックで単位固定されている人は移動禁止。
      const lock = state.lockType[staff.index][dayIndex];
      if (lock === '第1固定' || lock === '第2固定') {
        return;
      }

      moveFromOtherUnit = true;
    }

    let priority = unitPreferenceScore_(ctx, staff, dayIndex, unit, state);

    // 元所属を最優先
    if (staff.unit === unit) priority += 10000;

    // 共通所属は次点
    if (!staff.unit || staff.unit === '共通') priority += 7000;

    // パート活用を社員ヘルプより優先
    if (staff.employment === 'パート') priority += 4000;
    if (staff.employment === 'ドライバー') priority -= 3000;

    // 既に勤務中なら勤務日数を増やさず済む
    if (isWorking) priority += 3000;

    // 希望休は可能な限り避ける
    if (staff.preferences[dayIndex] === '希望休') priority -= 50000;

    priority += numberOrZero_(staff.scores[dayIndex]);

    candidates.push({
      staff: staff,
      priority: priority,
      moveFromOtherUnit: moveFromOtherUnit
    });
  });

  candidates.sort(function (a, b) {
    if (b.priority !== a.priority) return b.priority - a.priority;
    return a.staff.index - b.staff.index;
  });

  return candidates.length > 0 ? candidates[0] : null;
}


/* =========================================================
 * 13. 送迎
 * ========================================================= */

function ensureTransportStaffing_(ctx, state) {
  if (ctx.transportUse !== '使用する') {
    return;
  }

  for (let d = 0; d < ctx.daysInMonth; d++) {
    const day = ctx.days[d];
    const needed = getTransportNeedForDay_(ctx, day);
    if (needed <= 0) continue;

    const isOperating = day.state !== '休日' &&
      (day.usage1 > 0 || day.usage2 > 0 ||
       day.legal1 > 0 || day.legal2 > 0);

    if (!isOperating) {
      continue;
    }

    // まず運転者を確保
    ensureDailyCapabilityCount_(
      ctx,
      state,
      d,
      needed,
      function (staff) {
        return staff.canDrive;
      },
      '運転者'
    );

    // 社員同乗が「必須」の施設だけ、通常運用車両数ぶん確保する。
    // 「施設ルールによる」は初期設定検証で生成を止める。
    if (ctx.employeeRideRule === '必須') {
      ensureDailyCapabilityCount_(
        ctx,
        state,
        d,
        needed,
        function (staff) {
          return staff.employment === '社員' && staff.canTransport;
        },
        '社員同乗'
      );
    }
  }
}


function ensureDailyCapabilityCount_(
  ctx,
  state,
  dayIndex,
  needed,
  predicate,
  label
) {
  function currentCount() {
    let count = 0;

    ctx.staff.forEach(function (staff) {
      if (!staff.active) return;
      if (state.work[staff.index][dayIndex] !== 1) return;
      if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;
      if (predicate(staff)) count++;
    });

    return count;
  }

  let guard = 0;

  while (currentCount() < needed && guard < 50) {
    guard++;

    const candidates = ctx.staff.filter(function (staff) {
      if (!staff.active) return false;
      if (!predicate(staff)) return false;
      if (state.work[staff.index][dayIndex] === 1) return false;
      if (state.immutable[staff.index][dayIndex]) return false;
      if (hardUnavailableReason_(ctx, staff, dayIndex)) return false;
      if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return false;

      return !wouldExceedMaxConsecutive_(
        state.work[staff.index],
        dayIndex,
        getStaffMaxConsecutive_(ctx, staff, dayIndex)
      );
    });

    candidates.sort(function (a, b) {
      // 運転・社員同乗を同時に満たせる人を優先。
      const comboA =
        (a.canDrive ? 1 : 0) +
        (a.employment === '社員' && a.canTransport ? 1 : 0);
      const comboB =
        (b.canDrive ? 1 : 0) +
        (b.employment === '社員' && b.canTransport ? 1 : 0);

      if (comboB !== comboA) return comboB - comboA;

      const scoreA = candidateDayScore_(ctx, a, dayIndex, state);
      const scoreB = candidateDayScore_(ctx, b, dayIndex, state);

      if (scoreB !== scoreA) return scoreB - scoreA;
      return a.index - b.index;
    });

    if (candidates.length === 0) {
      break;
    }

    const staff = candidates[0];
    state.work[staff.index][dayIndex] = 1;

    // 単位未割当なら、その日の不足が大きい方へ。
    if (!getAssignedUnit_(ctx, state, staff.index, dayIndex)) {
      const deficit1 = Math.max(
        0,
        requiredDirectSupportCount_(ctx, ctx.days[dayIndex], '第1') -
        countLegalStaff_(ctx, state, dayIndex, '第1')
      );

      const deficit2 = ctx.unitCount >= 2
        ? Math.max(
            0,
            requiredDirectSupportCount_(ctx, ctx.days[dayIndex], '第2') -
            countLegalStaff_(ctx, state, dayIndex, '第2')
          )
        : -1;

      const fallback = ctx.unitCount === 1 || deficit1 >= deficit2 ? '第1' : '第2';
      const chosen = choosePreferredUnitForStaff_(ctx, state, staff, dayIndex, fallback);
      if (chosen) setAssignedUnit_(ctx, state, staff.index, dayIndex, chosen);
    }
  }

  if (currentCount() < needed) {
    state.errors.push(
      (dayIndex + 1) + '日：' + label +
      'が必要 ' + needed + '名に対して ' +
      currentCount() + '名しか確保できません。'
    );
  }
}


/* =========================================================
 * 14. 加配・専門支援（ソフト条件）
 * ========================================================= */

function applySoftRoleTargets_(ctx, state) {
  for (let d = 0; d < ctx.daysInMonth; d++) {
    const day = ctx.days[d];

    if (day.state === '休日') {
      continue;
    }

    for (let u = 1; u <= ctx.unitCount; u++) {
      const unit = '第' + u;
      const usage = u === 1 ? day.usage1 : day.usage2;

      if (usage <= 0) {
        continue;
      }

      const config = ctx.placement[unit];
      if (!config) continue;

      // 児童指導員等加配
      if (config.addOn &&
          config.addOn.operation === '算定する') {
        tryPlaceSoftNamedRole_(
          ctx,
          state,
          d,
          unit,
          config.addOn,
          '児童指導員等加配',
          true
        );
      }

      // 専門的支援実施
      if (config.specialExecution &&
          config.specialExecution.operation === '使用する') {
        tryPlaceSoftNamedRole_(
          ctx,
          state,
          d,
          unit,
          config.specialExecution,
          '専門支援実施',
          false
        );
      }

      // 専門的支援体制加算
      if (config.specialSystem &&
          config.specialSystem.operation === '算定する') {
        tryPlaceSoftNamedRole_(
          ctx,
          state,
          d,
          unit,
          config.specialSystem,
          '専門支援体制',
          true
        );
      }
    }
  }
}


function tryPlaceSoftNamedRole_(
  ctx,
  state,
  dayIndex,
  unit,
  config,
  roleLabel,
  mustBeExtra
) {
  const refs = [config.main, config.sub].filter(Boolean);

  // すでに担当者が同単位で勤務しているなら成立扱い。
  for (let r = 0; r < refs.length; r++) {
    const existing = findStaffByRef_(ctx, refs[r]);
    if (!existing) continue;

    if (state.work[existing.index][dayIndex] === 1 &&
        getAssignedUnit_(ctx, state, existing.index, dayIndex) === unit) {
      return true;
    }
  }

  for (let r = 0; r < refs.length; r++) {
    const staff = findStaffByRef_(ctx, refs[r]);

    if (!staff || !staff.active) continue;
    if (hardUnavailableReason_(ctx, staff, dayIndex)) continue;

    if (state.immutable[staff.index][dayIndex]) {
      if (state.work[staff.index][dayIndex] !== 1) {
        continue;
      }

      const fixedUnit = getAssignedUnit_(
        ctx,
        state,
        staff.index,
        dayIndex
      );

      if (fixedUnit !== unit) {
        continue;
      }
    }

    if (state.work[staff.index][dayIndex] !== 1 &&
        wouldExceedMaxConsecutive_(
          state.work[staff.index],
          dayIndex,
          getStaffMaxConsecutive_(ctx, staff, dayIndex)
        )) {
      continue;
    }

    // 追加配置扱いが必要な場合、担当者を法定人数から除外しても
    // 基準配置を満たせる見込みがあるか確認する。
    if (mustBeExtra && staff.legalEligible && staff.directSupport) {
      const needed = unit === '第1'
        ? ctx.days[dayIndex].legal1
        : ctx.days[dayIndex].legal2;

      const current = countLegalStaff_(
        ctx,
        state,
        dayIndex,
        unit
      );

      const isAlreadyCounted =
        state.work[staff.index][dayIndex] === 1 &&
        getAssignedUnit_(ctx, state, staff.index, dayIndex) === unit &&
        state.role[staff.index][dayIndex] !== '児発管:' + unit;

      const baselineAfterExcluding = current - (isAlreadyCounted ? 1 : 0);

      if (baselineAfterExcluding < needed) {
        // まず基準配置を追加で確保してから再判定。
        ensureLegalStaffing_(ctx, state);
      }
    }

    state.work[staff.index][dayIndex] = 1;
    setAssignedUnit_(ctx, state, staff.index, dayIndex, unit);

    if (!state.role[staff.index][dayIndex]) {
      state.role[staff.index][dayIndex] = roleLabel + ':' + unit;
    }

    return true;
  }

  if (refs.length > 0) {
    state.warnings.push(
      (dayIndex + 1) + '日 ' + unit + '：' +
      roleLabel + '担当を配置できませんでした。'
    );
  }

  return false;
}


function repairPairRules_(ctx, state) {
  const rows = ctx.rules && ctx.rules.pairRules || [];
  rows.forEach(function (r) {
    if (!isRuleEnabled_(r) || String(r['自動調整'] || '').trim() === '手動のみ') return;
    const a = findStaffByRef_(ctx, r['職員A']);
    const b = findStaffByRef_(ctx, r['職員B']);
    if (!a || !b) return;
    const type = String(r['ルール'] || '').trim();
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (!ruleDateMatches_(ctx, r, d)) continue;
      let aw = state.work[a.index][d] === 1;
      let bw = state.work[b.index][d] === 1;
      if (type === '合わせる' || type === '同時休み優先') {
        if (aw === bw) continue;
        const add = aw ? b : a;
        if (!state.immutable[add.index][d] && !hardUnavailableReason_(ctx, add, d) &&
            !wouldExceedMaxConsecutive_(state.work[add.index], d, getStaffMaxConsecutive_(ctx, add, d))) {
          state.work[add.index][d] = 1;
          const fallback = add.unit === '第1' || add.unit === '第2' ? add.unit : chooseUnitByDeficit_(ctx, state, d);
          const chosen = choosePreferredUnitForStaff_(ctx, state, add, d, fallback);
          if (chosen) setAssignedUnit_(ctx, state, add.index, d, chosen);
        }
      } else if (type === 'どちらか一方') {
        if (aw !== bw) continue;
        if (!aw && !bw) {
          const candidates = [a,b].filter(function (st) {
            return !state.immutable[st.index][d] && !hardUnavailableReason_(ctx, st, d) &&
              !wouldExceedMaxConsecutive_(state.work[st.index], d, getStaffMaxConsecutive_(ctx, st, d));
          }).sort(function (x,y) { return candidateDayScore_(ctx,y,d,state)-candidateDayScore_(ctx,x,d,state); });
          if (candidates.length) {
            const st = candidates[0];
            state.work[st.index][d] = 1;
            const fallback = st.unit === '第1' || st.unit === '第2' ? st.unit : chooseUnitByDeficit_(ctx, state, d);
            const chosen = choosePreferredUnitForStaff_(ctx, state, st, d, fallback);
            if (chosen) setAssignedUnit_(ctx, state, st.index, d, chosen);
          }
        } else {
          const candidates = [a,b].filter(function (st) { return canRemoveStaffDay_(ctx, state, st.index, d); })
            .sort(function (x,y) { return candidateDayScore_(ctx,x,d,state)-candidateDayScore_(ctx,y,d,state); });
          if (candidates.length) {
            const st = candidates[0];
            state.work[st.index][d] = 0; state.unit[st.index][d] = ''; state.role[st.index][d] = '';
          }
        }
      } else if (type === '合わせない') {
        if (!(aw && bw)) continue;
        const candidates = [a,b].filter(function (st) { return canRemoveStaffDay_(ctx, state, st.index, d); })
          .sort(function (x,y) { return candidateDayScore_(ctx,x,d,state)-candidateDayScore_(ctx,y,d,state); });
        if (candidates.length) {
          const st = candidates[0];
          state.work[st.index][d] = 0; state.unit[st.index][d] = ''; state.role[st.index][d] = '';
        }
      } else if (type === '同時休み禁止') {
        if (aw || bw) continue;
        const candidates = [a,b].filter(function (st) {
          return !state.immutable[st.index][d] && !hardUnavailableReason_(ctx, st, d) &&
            !wouldExceedMaxConsecutive_(state.work[st.index], d, getStaffMaxConsecutive_(ctx, st, d));
        }).sort(function (x,y) { return candidateDayScore_(ctx,y,d,state)-candidateDayScore_(ctx,x,d,state); });
        if (candidates.length) {
          const st = candidates[0];
          state.work[st.index][d] = 1;
          const fallback = st.unit === '第1' || st.unit === '第2' ? st.unit : chooseUnitByDeficit_(ctx, state, d);
          const chosen = choosePreferredUnitForStaff_(ctx, state, st, d, fallback);
          if (chosen) setAssignedUnit_(ctx, state, st.index, d, chosen);
        }
      } else if ((type === '同一Unit優先' || type === '別Unit優先') && aw && bw) {
        const au = getAssignedUnit_(ctx, state, a.index, d);
        const bu = getAssignedUnit_(ctx, state, b.index, d);
        const satisfied = type === '同一Unit優先' ? au === bu : au && bu && au !== bu;
        if (satisfied) continue;
        if (!state.immutable[b.index][d]) {
          let desired = '';
          if (type === '同一Unit優先') desired = au;
          else desired = au === '第1' ? '第2' : '第1';
          if (desired && isUnitAllowedForStaff_(ctx, b, d, desired)) setAssignedUnit_(ctx, state, b.index, d, desired);
        }
      }
    }
  });
}

function repairMinimumConsecutiveRest_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') return;
    const rules = getPersonRules_(ctx, staff, '最低連休', null);
    if (!rules.length) return;
    const needed = Math.max.apply(null, rules.map(function (r) { return numberOrZero_(r['値1']); }).filter(function (n) { return n > 0; }));
    if (!(needed > 0) || hasConsecutiveOff_(state.work[staff.index], ctx.daysInMonth, needed)) return;

    let best = null;
    for (let start = 0; start <= ctx.daysInMonth - needed; start++) {
      const inside = [];
      let blocked = false;
      let loss = 0;
      for (let d = start; d < start + needed; d++) {
        if (state.work[staff.index][d] !== 1) continue;
        if (!canRemoveStaffDay_(ctx, state, staff.index, d)) { blocked = true; break; }
        inside.push(d);
        loss += candidateDayScore_(ctx, staff, d, state);
      }
      if (blocked) continue;
      if (!best || inside.length < best.inside.length || (inside.length === best.inside.length && loss < best.loss)) {
        best = {start:start, inside:inside, loss:loss};
      }
    }
    if (!best) {
      state.warnings.push(staff.name + '：最低連休を自動確保できません。');
      return;
    }

    const moved = [];
    for (let i = 0; i < best.inside.length; i++) {
      const fromDay = best.inside[i];
      const alternatives = [];
      for (let d = 0; d < ctx.daysInMonth; d++) {
        if (d >= best.start && d < best.start + needed) continue;
        if (state.work[staff.index][d] === 1 || state.immutable[staff.index][d]) continue;
        if (hardUnavailableReason_(ctx, staff, d)) continue;
        const test = state.work[staff.index].slice();
        test[fromDay] = 0;
        if (wouldExceedMaxConsecutive_(test, d, getStaffMaxConsecutive_(ctx, staff, d))) continue;
        alternatives.push({day:d, score:candidateDayScore_(ctx,staff,d,state)});
      }
      alternatives.sort(function (a,b) { return b.score-a.score; });
      if (!alternatives.length) continue;
      const toDay = alternatives[0].day;
      const oldUnit = getAssignedUnit_(ctx, state, staff.index, fromDay) || staff.unit;
      state.work[staff.index][fromDay] = 0; state.unit[staff.index][fromDay] = ''; state.role[staff.index][fromDay] = '';
      state.work[staff.index][toDay] = 1;
      const fallback = oldUnit === '第1' || oldUnit === '第2' ? oldUnit : chooseUnitByDeficit_(ctx,state,toDay);
      const chosen = choosePreferredUnitForStaff_(ctx,state,staff,toDay,fallback);
      if (chosen) setAssignedUnit_(ctx,state,staff.index,toDay,chosen);
      moved.push(fromDay);
    }
    if (!hasConsecutiveOff_(state.work[staff.index], ctx.daysInMonth, needed)) {
      state.warnings.push(staff.name + '：最低連休' + needed + '日を自動確保できません。');
    }
  });
}

/* =========================================================
 * 15. 契約勤務日数
 * ========================================================= */

function normalizeContractDays_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') {
      return;
    }

    const target = getWorkTargetDays_(staff);
    let current = countWorkDays_(
      state.work[staff.index],
      ctx.daysInMonth
    );

    // 不足分を追加
    if (current < target) {
      const candidates = [];

      for (let d = 0; d < ctx.daysInMonth; d++) {
        if (state.work[staff.index][d] === 1) continue;
        if (state.immutable[staff.index][d]) continue;
        if (hardUnavailableReason_(ctx, staff, d)) continue;

        candidates.push({
          dayIndex: d,
          score: candidateDayScore_(ctx, staff, d, state)
        });
      }

      candidates.sort(function (a, b) {
        if (b.score !== a.score) return b.score - a.score;
        return a.dayIndex - b.dayIndex;
      });

      for (let i = 0; i < candidates.length && current < target; i++) {
        const d = candidates[i].dayIndex;

        if (wouldExceedMaxConsecutive_(
          state.work[staff.index],
          d,
          getStaffMaxConsecutive_(ctx, staff, d)
        )) {
          continue;
        }

        state.work[staff.index][d] = 1;

        if (!getAssignedUnit_(ctx, state, staff.index, d)) {
          const baseUnit = staff.unit === '第1' || staff.unit === '第2'
            ? staff.unit
            : chooseUnitByDeficit_(ctx, state, d);
          const preferred = choosePreferredUnitForStaff_(ctx, state, staff, d, baseUnit);

          setAssignedUnit_(
            ctx,
            state,
            staff.index,
            d,
            preferred
          );
        }

        current++;
      }
    }

    // 超過分を、安全に外せる日から削る
    if (current > target) {
      const removable = [];

      for (let d = 0; d < ctx.daysInMonth; d++) {
        if (state.work[staff.index][d] !== 1) continue;
        if (state.immutable[staff.index][d]) continue;
        if (state.role[staff.index][d].indexOf('児発管:') === 0) continue;

        removable.push({
          dayIndex: d,
          score: candidateDayScore_(ctx, staff, d, state)
        });
      }

      // スコアが低い日から外す。
      removable.sort(function (a, b) {
        if (a.score !== b.score) return a.score - b.score;
        return b.dayIndex - a.dayIndex;
      });

      for (let i = 0; i < removable.length && current > target; i++) {
        const d = removable[i].dayIndex;

        if (!canRemoveStaffDay_(ctx, state, staff.index, d)) {
          continue;
        }

        state.work[staff.index][d] = 0;
        state.unit[staff.index][d] = '';
        state.role[staff.index][d] = '';
        current--;
      }
    }

    if (current !== target) {
      state.warnings.push(
        staff.name + '：必要勤務 ' + target +
        '日に対して生成結果が ' + current + '日です。'
      );
    }
  });
}


function canRemoveStaffDay_(ctx, state, staffIndex, dayIndex) {
  const staff = ctx.staff[staffIndex];

  if (state.work[staffIndex][dayIndex] !== 1) {
    return true;
  }

  if (isAutoAdjustBlocked_(ctx, staff, dayIndex)) {
    return false;
  }

  if (state.immutable[staffIndex][dayIndex]) {
    return false;
  }

  if (state.role[staffIndex][dayIndex].indexOf('児発管:') === 0) {
    return false;
  }

  const unit = getAssignedUnit_(ctx, state, staffIndex, dayIndex);

  if (unit === '第1' || unit === '第2') {
    if (staff.directSupport &&
        staff.legalEligible &&
        staff.halfDays.indexOf(dayIndex + 1) < 0) {
      const needed = requiredDirectSupportCount_(ctx, ctx.days[dayIndex], unit);

      const current = countLegalStaff_(
        ctx,
        state,
        dayIndex,
        unit
      );

      if (current - 1 < needed) {
        return false;
      }
    }
  }

  if (ctx.transportUse === '使用する') {
    const needed = getTransportNeedForDay_(ctx, ctx.days[dayIndex]);
    const day = ctx.days[dayIndex];
    const isOperating = day.state !== '休日' &&
      (day.usage1 > 0 || day.usage2 > 0 ||
       day.legal1 > 0 || day.legal2 > 0);

    if (isOperating && needed > 0) {
      if (staff.canDrive &&
          countDailyCapability_(
            ctx,
            state,
            dayIndex,
            function (x) { return x.canDrive; }
          ) - 1 < needed) {
        return false;
      }

      if (ctx.employeeRideRule === '必須' &&
          staff.employment === '社員' &&
          staff.canTransport &&
          countDailyCapability_(
            ctx,
            state,
            dayIndex,
            function (x) {
              return x.employment === '社員' && x.canTransport;
            }
          ) - 1 < needed) {
        return false;
      }
    }
  }

  return true;
}


function countDailyCapability_(ctx, state, dayIndex, predicate) {
  let count = 0;

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    if (state.work[staff.index][dayIndex] !== 1) return;
    if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;
    if (predicate(staff)) count++;
  });

  return count;
}



/**
 * 勤務日数を変えず、勤務日を入れ替えて契約時間へ近づける。
 * 勤務時刻そのものは変更しないため、v1.0で安全にできる範囲の補正。
 */
function repairContractHoursByDaySwap_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') return;
    if (!(staff.minimumHours > 0)) return;

    const lower = staff.minimumHours;
    const upper = staff.targetHours > lower ? staff.targetHours : lower + 10;
    let hours = estimateStaffHours_(ctx, state, staff);
    if (hours >= lower && hours <= upper) return;

    const working = [];
    const off = [];
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (state.immutable[staff.index][d]) continue;
      if (state.work[staff.index][d] === 1) working.push(d);
      else if (!hardUnavailableReason_(ctx, staff, d)) off.push(d);
    }

    let best = null;
    working.forEach(function (fromDay) {
      if (!canRemoveStaffDay_(ctx, state, staff.index, fromDay)) return;
      off.forEach(function (toDay) {
        if (wouldExceedMaxConsecutive_(state.work[staff.index], toDay, getStaffMaxConsecutive_(ctx, staff, toDay))) return;
        const before = dayHoursForStaff_(ctx, staff, fromDay);
        const after = dayHoursForStaff_(ctx, staff, toDay);
        const next = hours - before + after;
        const currentDistance = distanceToRange_(hours, lower, upper);
        const nextDistance = distanceToRange_(next, lower, upper);
        if (nextDistance >= currentDistance) return;
        const score = nextDistance * 1000 - candidateDayScore_(ctx, staff, toDay, state);
        if (!best || score < best.score) {
          best = {fromDay: fromDay, toDay: toDay, nextHours: next, score: score};
        }
      });
    });

    if (best) {
      const oldUnit = getAssignedUnit_(ctx, state, staff.index, best.fromDay) || staff.unit;
      state.work[staff.index][best.fromDay] = 0;
      state.unit[staff.index][best.fromDay] = '';
      state.role[staff.index][best.fromDay] = '';
      state.work[staff.index][best.toDay] = 1;
      setAssignedUnit_(ctx, state, staff.index, best.toDay,
        choosePreferredUnitForStaff_(ctx, state, staff, best.toDay,
          oldUnit === '第1' || oldUnit === '第2' ? oldUnit : chooseUnitByDeficit_(ctx, state, best.toDay)));
      state.warnings.push(
        staff.name + '：契約時間調整のため ' + (best.fromDay + 1) + '日→' +
        (best.toDay + 1) + '日に勤務日を入替しました。'
      );
    }
  });
}

function estimateStaffHours_(ctx, state, staff) {
  let total = 0;
  for (let d = 0; d < ctx.daysInMonth; d++) {
    if (state.work[staff.index][d] === 1) total += dayHoursForStaff_(ctx, staff, d);
  }
  if (staff.paidLeaveDays > 0 && staff.requiredDays > 0 && staff.minimumHours > 0) {
    total += staff.paidLeaveDays * (staff.minimumHours / staff.requiredDays);
  }
  return total;
}

function dayHoursForStaff_(ctx, staff, dayIndex) {
  const advancedPattern = getPersonPatternRule_(ctx, staff, dayIndex);
  if (advancedPattern) {
    return getPatternHours_(ctx, advancedPattern, 8);
  }
  const dayNo = dayIndex + 1;
  if (staff.halfDays.indexOf(dayNo) >= 0) {
    return getPatternHours_(ctx, '単半休日', 5);
  }
  if (ctx.days[dayIndex] && ctx.days[dayIndex].dayType === '学校休日') {
    return getPatternHours_(ctx, '学校休日通常', getPatternHours_(ctx, '平日通常', 8));
  }
  return getPatternHours_(ctx, '平日通常', 8);
}

function getPatternHours_(ctx, patternName, fallback) {
  const rows = ctx.patternValues || [];
  if (!rows.length) return fallback;
  const headers = (rows[0] || []).map(String);
  const header = buildHeaderMap_(headers);
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const name = String(valueAt_(row, header, '勤務パターン名') || valueAt_(row, header, 'パターン名') || row[1] || '').trim();
    if (name !== patternName) continue;
    const h = Number(valueAt_(row, header, '実働時間') || row[7]);
    if (Number.isFinite(h) && h > 0) return h;
  }
  return fallback;
}

function distanceToRange_(value, lower, upper) {
  if (value < lower) return lower - value;
  if (value > upper) return value - upper;
  return 0;
}

/* =========================================================
 * 16. 連勤
 * ========================================================= */

function repairMaxConsecutive_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') {
      return;
    }

    const maxCap = getStaffMaxConsecutive_(ctx, staff);
    let guard = 0;

    while (getMaxConsecutive_(
      state.work[staff.index],
      ctx.daysInMonth
    ) > maxCap && guard < 40) {
      guard++;

      const run = findFirstExcessRun_(
        state.work[staff.index],
        ctx.daysInMonth,
        maxCap
      );

      if (!run) break;

      let repaired = false;

      // 超過連勤の中で、最も外しやすい日を探す。
      const runDays = [];

      for (let d = run.start; d <= run.end; d++) {
        if (state.immutable[staff.index][d]) continue;
        if (isAutoAdjustBlocked_(ctx, staff, d)) continue;
        if (state.role[staff.index][d].indexOf('児発管:') === 0) continue;

        runDays.push({
          dayIndex: d,
          score: candidateDayScore_(ctx, staff, d, state)
        });
      }

      runDays.sort(function (a, b) {
        return a.score - b.score;
      });

      for (let r = 0; r < runDays.length && !repaired; r++) {
        const removeDay = runDays[r].dayIndex;

        if (!canRemoveStaffDay_(ctx, state, staff.index, removeDay)) {
          continue;
        }

        // 代替出勤日を探す。
        const alternatives = [];

        for (let d = 0; d < ctx.daysInMonth; d++) {
          if (state.work[staff.index][d] === 1) continue;
          if (state.immutable[staff.index][d]) continue;
          if (hardUnavailableReason_(ctx, staff, d)) continue;

          const testRow = state.work[staff.index].slice();
          testRow[removeDay] = 0;

          if (wouldExceedMaxConsecutive_(
            testRow,
            d,
            maxCap
          )) {
            continue;
          }

          alternatives.push({
            dayIndex: d,
            score: candidateDayScore_(ctx, staff, d, state)
          });
        }

        alternatives.sort(function (a, b) {
          return b.score - a.score;
        });

        if (alternatives.length === 0) {
          continue;
        }

        const addDay = alternatives[0].dayIndex;

        state.work[staff.index][removeDay] = 0;
        state.unit[staff.index][removeDay] = '';
        state.role[staff.index][removeDay] = '';

        state.work[staff.index][addDay] = 1;

        const baseUnit = staff.unit === '第1' || staff.unit === '第2'
          ? staff.unit
          : chooseUnitByDeficit_(ctx, state, addDay);
        const preferred = choosePreferredUnitForStaff_(ctx, state, staff, addDay, baseUnit);

        setAssignedUnit_(
          ctx,
          state,
          staff.index,
          addDay,
          preferred
        );

        repaired = true;
      }

      if (!repaired) {
        state.errors.push(
          staff.name + '：最大' + maxCap +
          '連勤を超える配置を解消できません。'
        );
        break;
      }
    }
  });
}


function improvePreferredConsecutive_(ctx, state) {
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') return;

    const preferredCap = getStaffPreferredConsecutive_(ctx, staff);
    const maxCap = getStaffMaxConsecutive_(ctx, staff);
    if (preferredCap >= maxCap) return;
    const maxNow = getMaxConsecutive_(
      state.work[staff.index],
      ctx.daysInMonth
    );

    if (maxNow <= preferredCap) {
      return;
    }

    // ここでは無理な入替をせず、最大連勤以下なら安全条件を優先する。
    // v1.0では警告として残し、運用余裕不足として人間が確認できるようにする。
    state.warnings.push(
      staff.name + '：原則' + preferredCap +
      '連勤を超える連続勤務があります（最大' +
      maxCap + '連勤以内）。'
    );
  });
}


/* =========================================================
 * 17. 出力
 * ========================================================= */

function writePlan_(ctx, state) {
  const autoSheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.AUTO);

  const workOutput = matrix_(
    SHIFT_APP.MAX_STAFF,
    SHIFT_APP.MAX_DAYS,
    0
  );

  const unitOutput = matrix_(
    SHIFT_APP.MAX_STAFF,
    SHIFT_APP.MAX_DAYS,
    ''
  );

  for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
    const staff = ctx.staff[s];

    for (let d = 0; d < SHIFT_APP.MAX_DAYS; d++) {
      if (!staff.active || d >= ctx.daysInMonth) {
        workOutput[s][d] = 0;
        unitOutput[s][d] = '';
        continue;
      }

      workOutput[s][d] = state.work[s][d] ? 1 : 0;

      if (!state.work[s][d]) {
        unitOutput[s][d] = '';
        continue;
      }

      const assigned = getAssignedUnit_(ctx, state, s, d);

      // 元所属と同じなら空白にして既存数式の「所属」を使う。
      if (assigned &&
          (staff.unit !== assigned ||
           (staff.unit !== '第1' && staff.unit !== '第2'))) {
        unitOutput[s][d] = assigned;
      } else {
        unitOutput[s][d] = '';
      }
    }
  }

  autoSheet
    .getRange(SHIFT_APP.RANGE.WORK_PLAN)
    .setValues(workOutput);

  autoSheet
    .getRange(SHIFT_APP.RANGE.UNIT_PLAN)
    .setValues(unitOutput);

  appendLog_(
    'INFO',
    'シフト出力',
    Utilities.formatDate(
      ctx.targetMonth,
      ctx.ss.getSpreadsheetTimeZone(),
      'yyyy-MM'
    ),
    '内部勤務グリッドを確定しました。'
  );
}


/* =========================================================
 * 17.5 表示補助・配置理由・実績
 * ========================================================= */

function writeOperationStatus_(ss, status, nextAction, detail) {
  const sh = ss.getSheetByName(SHIFT_APP.SHEETS.CONTROL);
  if (!sh) return;
  setLabelValue_(sh, '運用ステータス', status || '');
  setLabelValue_(sh, '次にやること', nextAction || '');
  setLabelValue_(sh, 'ステータス詳細', detail || '');
  setLabelValue_(sh, '最終更新', new Date());
  const cells = sh.createTextFinder('最終更新').matchEntireCell(true).findAll();
  if (cells.length) cells[0].offset(0,1).setNumberFormat('yyyy/m/d HH:mm');
}

function ensurePlacementReasonSheet_(ss) {
  let sh = ss.getSheetByName(SHIFT_APP.SHEETS.REASON);
  if (!sh) sh = ss.insertSheet(SHIFT_APP.SHEETS.REASON);
  sh.getRange('A5:H600').clearContent();
  sh.getRange('A1').setValue('ウッディーのウキウキシフト｜配置理由');
  sh.getRange('A2').setValue('自動生成された配置について、主な理由を確認するための一覧です。');
  sh.getRange('A4:H4').setValues([['日付','曜日','職員ID','氏名','Unit','主な理由','適用ルール','候補スコア']]);
  sh.setFrozenRows(4);
  return sh;
}

function writePlacementReasons_(ctx, state) {
  const sh = ensurePlacementReasonSheet_(ctx.ss);
  const rows = [];
  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (state.work[staff.index][d] !== 1) continue;
      const reasons = [];
      const ruleLabels = [];
      const role = String(state.role[staff.index][d] || '').trim();
      const lock = String(state.lockType[staff.index][d] || '').trim();
      if (role) reasons.push(role);
      if (lock) reasons.push(lock);
      if (staff.preferences[d] === '出勤希望') reasons.push('出勤希望');
      getPersonRules_(ctx, staff, null, d).forEach(function (r) {
        const t = String(r['ルール種別'] || '').trim();
        if (t) ruleLabels.push(t + (r['値1'] ? ':' + r['値1'] : ''));
      });
      if (!reasons.length) reasons.push('法定配置・契約・公平性を考慮した自動配置');
      rows.push([
        ctx.days[d].date,
        ctx.days[d].weekday,
        staff.id,
        staff.name,
        getAssignedUnit_(ctx, state, staff.index, d) || staff.unit || '',
        unique_(reasons).join(' / '),
        unique_(ruleLabels).join(' / '),
        Math.round(candidateDayScore_(ctx, staff, d, state))
      ]);
    }
  });
  if (rows.length) sh.getRange(5,1,rows.length,8).setValues(rows);
  sh.getRange('A5:A600').setNumberFormat('yyyy/m/d');
}

function ensureActualSheet_(ss) {
  let sh = ss.getSheetByName(SHIFT_APP.SHEETS.ACTUAL);
  if (!sh) sh = ss.insertSheet(SHIFT_APP.SHEETS.ACTUAL);
  sh.getRange('A1').setValue('ウッディーのウキウキシフト｜実績');
  sh.getRange('A2').setValue('確定後、実績区分・開始・終了を入力すると予定との差異を確認できます。');
  sh.getRange('A4:M4').setValues([['日付','曜日','職員ID','氏名','予定区分','予定Unit','予定時間','実績区分','実績開始','実績終了','実績時間','差異時間','メモ']]);
  sh.setFrozenRows(4);
  return sh;
}

function initializeActualSheet_(ctx, plan) {
  const sh = ensureActualSheet_(ctx.ss);
  const oldValues = sh.getRange('A5:M600').getValues();
  const oldMap = {};
  oldValues.forEach(function (r) {
    const date = normalizeDate_(r[0]);
    const id = String(r[2] || '').trim();
    if (!date || !id) return;
    const key = Utilities.formatDate(date, ctx.ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') + '|' + id;
    oldMap[key] = r.slice(7,13);
  });

  sh.getRange('A5:M600').clearContent();
  const rows = [];
  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    for (let d = 0; d < ctx.daysInMonth; d++) {
      const works = plan.work[staff.index][d] === 1;
      const key = Utilities.formatDate(ctx.days[d].date, ctx.ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') + '|' + staff.id;
      const old = oldMap[key] || ['', '', '', '', '', ''];
      rows.push([
        ctx.days[d].date,
        ctx.days[d].weekday,
        staff.id,
        staff.name,
        works ? '勤務' : '休み',
        works ? (getAssignedUnitFromMatrices_(staff, plan.unit, staff.index, d) || staff.unit || '') : '',
        works ? dayHoursForStaff_(ctx, staff, d) : 0,
        old[0] || '', old[1] || '', old[2] || '', '', '', old[5] || ''
      ]);
    }
  });
  if (!rows.length) return;
  sh.getRange(5,1,rows.length,13).setValues(rows);

  const actualHourFormulas = [];
  const diffFormulas = [];
  for (let i = 0; i < rows.length; i++) {
    const r = 5 + i;
    actualHourFormulas.push(['=IF(OR(I'+r+'="",J'+r+'=""),"",MOD(J'+r+'-I'+r+',1)*24)']);
    diffFormulas.push(['=IF(K'+r+'="","",K'+r+'-G'+r+')']);
  }
  sh.getRange(5,11,rows.length,1).setFormulas(actualHourFormulas);
  sh.getRange(5,12,rows.length,1).setFormulas(diffFormulas);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['勤務','休み','有休','欠勤','遅刻','早退'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(5,8,rows.length,1).setDataValidation(rule);
  sh.getRange(5,1,rows.length,1).setNumberFormat('yyyy/m/d');
  sh.getRange(5,9,rows.length,2).setNumberFormat('hh:mm');
  sh.getRange(5,7,rows.length,1).setNumberFormat('0.00');
  sh.getRange(5,11,rows.length,2).setNumberFormat('0.00');
}

/* =========================================================
 * 18. 現在シフト検証
 * ========================================================= */

function validateCurrentShift_(ctx, plan) {
  SpreadsheetApp.flush();

  const errors = [];
  const warnings = [];

  const adminCheck = validateTokyoAdminCompliance_(ctx);
  Array.prototype.push.apply(errors, adminCheck.errors);
  Array.prototype.push.apply(warnings, adminCheck.warnings);

  const shiftSheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.SHIFT);
  const monthlySheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.MONTHLY);

  const shortage1 = shiftSheet.getRange('C7:AG7').getValues()[0];
  const shortage2 = shiftSheet.getRange('C10:AG10').getValues()[0];

  for (let d = 0; d < ctx.daysInMonth; d++) {
    if (ctx.days[d].state === '休日') continue;
    if (numberOrZero_(shortage1[d]) > 0) {
      errors.push(
        (d + 1) + '日 第1：法定配置が ' +
        numberOrZero_(shortage1[d]) + '名不足しています。'
      );
    }

    if (ctx.unitCount >= 2 &&
        numberOrZero_(shortage2[d]) > 0) {
      errors.push(
        (d + 1) + '日 第2：法定配置が ' +
        numberOrZero_(shortage2[d]) + '名不足しています。'
      );
    }

    if (ctx.transportUse === '使用する') {
      const needVehicles = getTransportNeedForDay_(ctx, ctx.days[d]);
      if (needVehicles > 0) {
        const driverCount = countDailyCapabilityFromPlan_(ctx, plan, d, function (staff) { return staff.canDrive; });
        if (driverCount < needVehicles) {
          errors.push((d + 1) + '日：送迎の運転者が必要' + needVehicles + '名に対して' + driverCount + '名です。');
        }
        if (ctx.employeeRideRule === '必須') {
          const employeeRideCount = countDailyCapabilityFromPlan_(ctx, plan, d, function (staff) {
            return staff.employment === '社員' && staff.canTransport;
          });
          if (employeeRideCount < needVehicles) {
            errors.push((d + 1) + '日：社員同乗が必要' + needVehicles + '名に対して' + employeeRideCount + '名です。');
          }
        }
      }
    }
  }

  // 月間集計：DAYS/HOURS
  const monthlyValues = monthlySheet.getRange('A3:P18').getDisplayValues();
  const headers = monthlyValues[0].map(String);
  const header = buildHeaderMap_(headers);

  for (let i = 1; i < monthlyValues.length; i++) {
    const row = monthlyValues[i];
    const id = String(valueAt_(row, header, '職員ID') || '').trim();
    const name = String(valueAt_(row, header, '氏名') || '').trim();
    const employment = String(valueAt_(row, header, '雇用区分') || '').trim();
    const dayJudge = String(valueAt_(row, header, '勤務日数判定') || '').trim();
    const hourJudge = String(valueAt_(row, header, '勤務時間判定') || '').trim();
    const total = String(valueAt_(row, header, '総合判定') || '').trim();

    if (!id || !name) continue;

    if (employment === '社員') {
      if (dayJudge && dayJudge !== 'OK') {
        errors.push(id + ' ' + name + '：契約勤務日数「' + dayJudge + '」');
      }
      if (hourJudge && hourJudge !== 'OK') {
        errors.push(id + ' ' + name + '：契約勤務時間「' + hourJudge + '」');
      }
      if (total && total !== 'OK') {
        errors.push(id + ' ' + name + '：月間集計「' + total + '」');
      }
    } else if (total && total !== 'OK' && total !== '対象外' && total !== '時間要確認') {
      warnings.push(id + ' ' + name + '：月間集計「' + total + '」');
    }
  }

  // 最大連勤をGAS側でも独立検証
  ctx.staff.forEach(function (staff) {
    if (!staff.active || staff.employment !== '社員') return;

    const maxRun = getMaxConsecutive_(
      plan.work[staff.index],
      ctx.daysInMonth
    );

    const maxCap = getStaffMaxConsecutive_(ctx, staff);
    const preferredCap = getStaffPreferredConsecutive_(ctx, staff);
    if (maxRun > maxCap) {
      errors.push(
        staff.name + '：最大連勤 ' + maxRun +
        '日（上限 ' + maxCap + '日）'
      );
    } else if (maxRun > preferredCap) {
      warnings.push(
        staff.name + '：原則' + preferredCap +
        '連勤を超え、最大 ' + maxRun + '連勤です。'
      );
    }

    const requestedOffViolations = [];

    for (let d = 0; d < ctx.daysInMonth; d++) {
      if (staff.preferences[d] === '希望休' &&
          plan.work[staff.index][d] === 1) {
        requestedOffViolations.push(d + 1);
      }
    }

    if (requestedOffViolations.length > 0) {
      warnings.push(
        staff.name + '：希望休に勤務 ' +
        requestedOffViolations.join(', ') + '日'
      );
    }
  });

  // 配置ロックの実現状況
  ctx.locks.forEach(function (lock) {
    if (lock.invalid) return;

    const staff = ctx.staff.find(function (x) {
      return x.id === lock.staffId && x.active;
    });
    if (!staff) return;

    const d = lock.day - 1;
    const works = plan.work[staff.index][d] === 1;
    const assigned = getAssignedUnitFromMatrices_(
      staff,
      plan.unit,
      staff.index,
      d
    );

    if (lock.type === '休み固定' && works) {
      errors.push(
        lock.day + '日 ' + staff.name +
        '：休み固定が反映されていません。'
      );
    }

    if (lock.type === '出勤固定' && !works) {
      errors.push(
        lock.day + '日 ' + staff.name +
        '：出勤固定が反映されていません。'
      );
    }

    if (lock.type === '第1固定' &&
        (!works || assigned !== '第1')) {
      errors.push(
        lock.day + '日 ' + staff.name +
        '：第1固定が反映されていません。'
      );
    }

    if (lock.type === '第2固定' &&
        (!works || assigned !== '第2')) {
      errors.push(
        lock.day + '日 ' + staff.name +
        '：第2固定が反映されていません。'
      );
    }
  });

  const advanced = validateAdvancedPlan_(ctx, plan);
  Array.prototype.push.apply(errors, advanced.errors);
  Array.prototype.push.apply(warnings, advanced.warnings);

  return {
    errors: unique_(errors),
    warnings: unique_(warnings)
  };
}


/* =========================================================
 * 19. 修正候補
 * ========================================================= */

function buildCorrectionSuggestions_(ctx, plan, validation, limit) {
  const suggestions = [];

  for (let d = 0; d < ctx.daysInMonth && suggestions.length < limit; d++) {
    const day = ctx.days[d];

    for (let u = 1; u <= ctx.unitCount && suggestions.length < limit; u++) {
      const unit = '第' + u;
      const needed = requiredDirectSupportCount_(ctx, day, unit);

      if (needed <= 0) continue;

      const current = countLegalStaffFromPlan_(
        ctx,
        plan,
        d,
        unit
      );

      if (current >= needed) continue;

      const candidates = ctx.staff.filter(function (staff) {
        if (!staff.active) return false;
        if (!staff.directSupport || !staff.legalEligible) return false;
        if (plan.work[staff.index][d] === 1) return false;
        if (hardUnavailableReason_(ctx, staff, d)) return false;
        return true;
      });

      if (candidates.length > 0) {
        suggestions.push(
          (d + 1) + '日 ' + unit + '：' +
          candidates.slice(0, 3).map(function (x) {
            return x.name;
          }).join(' / ') +
          ' を追加配置候補として確認'
        );
      }
    }
  }

  if (ctx.transportUse === '使用する' &&
      suggestions.length < limit) {
    for (let d = 0; d < ctx.daysInMonth && suggestions.length < limit; d++) {
      const day = ctx.days[d];
      const isOperating = day.state !== '休日' &&
        (day.usage1 > 0 || day.usage2 > 0 ||
         day.legal1 > 0 || day.legal2 > 0);

      if (!isOperating) continue;

      const driverCount = countDailyCapabilityFromPlan_(
        ctx,
        plan,
        d,
        function (staff) { return staff.canDrive; }
      );

      if (driverCount < getTransportNeedForDay_(ctx, day)) {
        const names = ctx.staff.filter(function (staff) {
          return staff.active &&
            staff.canDrive &&
            !plan.work[staff.index][d] &&
            !hardUnavailableReason_(ctx, staff, d);
        }).slice(0, 3).map(function (x) {
          return x.name;
        });

        if (names.length > 0) {
          suggestions.push(
            (d + 1) + '日：運転者候補 ' + names.join(' / ')
          );
        }
      }
    }
  }

  // 契約日数・時間の調整候補
  if (suggestions.length < limit) {
    const monthlySheet = ctx.ss.getSheetByName(SHIFT_APP.SHEETS.MONTHLY);
    if (monthlySheet) {
      const values = monthlySheet.getRange('A3:P18').getDisplayValues();
      const headers = (values[0] || []).map(String);
      const header = buildHeaderMap_(headers);
      for (let i = 1; i < values.length && suggestions.length < limit; i++) {
        const row = values[i] || [];
        const id = String(valueAt_(row, header, '職員ID') || '').trim();
        const name = String(valueAt_(row, header, '氏名') || '').trim();
        const emp = String(valueAt_(row, header, '雇用区分') || '').trim();
        const dayJudge = String(valueAt_(row, header, '勤務日数判定') || '').trim();
        const hourJudge = String(valueAt_(row, header, '勤務時間判定') || '').trim();
        if (!id || !name || emp !== '社員') continue;
        if (dayJudge && dayJudge !== 'OK') {
          suggestions.push(id + ' ' + name + '：契約日数が「' + dayJudge + '」。指定職員のみ再生成で日数調整候補を再計算');
          if (suggestions.length >= limit) break;
        }
        if (hourJudge && hourJudge !== 'OK') {
          suggestions.push(id + ' ' + name + '：契約時間が「' + hourJudge + '」。勤務日数を維持した入替補正または勤務パターン確認');
        }
      }
    }
  }

  // 配置系加算の取りこぼし候補
  if (suggestions.length < limit) {
    const revenueSheet = ctx.ss.getSheetByName(SHIFT_APP.SHEETS.REVENUE);
    if (revenueSheet) {
      const daily = revenueSheet.getRange('A43:N73').getDisplayValues();
      for (let i = 0; i < Math.min(ctx.daysInMonth, daily.length) && suggestions.length < limit; i++) {
        const row = daily[i] || [];
        const verdict = String(row[13] || '');
        if (verdict.indexOf('🟠') !== 0) continue;
        const misses = [];
        if (row[4] === '×') misses.push('第1加配');
        if (row[5] === '×') misses.push('第2加配');
        if (row[6] === '×') misses.push('第1専門体制');
        if (row[7] === '×') misses.push('第2専門体制');
        suggestions.push((i + 1) + '日：' + (misses.length ? misses.join(' / ') : '配置系加算') + ' の取りこぼし候補。法定配置を崩さず有資格者の再配置を確認');
      }
    }
  }

  // ハード条件を満たし、原則連勤などのソフト警告だけが残る場合
  if (suggestions.length === 0 &&
      validation.errors.length === 0 &&
      validation.warnings.length > 0) {
    const warningText = validation.warnings.slice(0, 3).join(' / ');
    suggestions.push(
      '注意のみ：' + warningText +
      '。最大連勤などのハード条件内であれば確定可能です。必要な場合のみ再最適化してください。'
    );
  }

  return suggestions.slice(0, limit);
}


function countLegalStaffFromPlan_(ctx, plan, dayIndex, unit) {
  let count = 0;

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    if (plan.work[staff.index][dayIndex] !== 1) return;
    if (!staff.directSupport || !staff.legalEligible) return;
    if (isWoodyTokyoMunicipality_(ctx) && !isTokyoCoreLegalStaff_(staff)) return;
    if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;

    const assigned = getAssignedUnitFromMatrices_(
      staff,
      plan.unit,
      staff.index,
      dayIndex
    );

    if (assigned === unit) count++;
  });

  return count;
}


function countDailyCapabilityFromPlan_(
  ctx,
  plan,
  dayIndex,
  predicate
) {
  let count = 0;

  ctx.staff.forEach(function (staff) {
    if (!staff.active) return;
    if (plan.work[staff.index][dayIndex] !== 1) return;
    if (staff.halfDays.indexOf(dayIndex + 1) >= 0) return;
    if (predicate(staff)) count++;
  });

  return count;
}



function writeCorrectionSuggestions_(ctx, suggestions, validation) {
  const sheet = ctx.ss.getSheetByName(SHIFT_APP.SHEETS.SUGGESTIONS);
  if (!sheet) return;
  const maxRows = 10;
  sheet.getRange('A5:I14').clearContent();
  const rows = [];
  const items = suggestions.slice(0, maxRows);
  items.forEach(function (text, i) {
    let category = '改善';
    let priority = '2';
    let action = '内容を確認し、必要なら部分再生成または配置ロックで調整';
    if (/法定|送迎|契約|最大連勤|ロック/.test(text)) {
      category = '安全'; priority = '1';
    } else if (/加配|専門|加算/.test(text)) {
      category = '加算'; priority = '3';
    } else if (/注意のみ|原則[0-9０-９]+連勤/.test(text)) {
      category = '注意'; priority = '2';
      action = '最大連勤などのハード条件内なら確定可。必要な場合のみ再最適化';
    }
    rows.push([priority, category, '－', text, '', '', '', action, '確認']);
  });
  if (!rows.length) {
    rows.push(['－','完了','－','重大な修正候補はありません','','','','現在の条件で再確認','OK']);
  }
  sheet.getRange(5,1,rows.length,9).setValues(rows);
}

function writeChangeComparison_(ctx, beforePlan, afterPlan, actionName) {
  const sheet = ctx.ss.getSheetByName(SHIFT_APP.SHEETS.CHANGE_DIFF);
  if (!sheet) return;
  sheet.getRange('A5:I200').clearContent();
  const rows = [];
  for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
    const staff = ctx.staff[s];
    if (!staff || !staff.active) continue;
    for (let d = 0; d < ctx.daysInMonth; d++) {
      const bw = beforePlan.work[s][d] ? 1 : 0;
      const aw = afterPlan.work[s][d] ? 1 : 0;
      const bu = getAssignedUnitFromMatrices_(staff, beforePlan.unit, s, d) || '';
      const au = getAssignedUnitFromMatrices_(staff, afterPlan.unit, s, d) || '';
      if (bw === aw && bu === au) continue;
      let kind = '所属変更';
      if (bw !== aw) kind = aw ? '出勤追加' : '休みへ変更';
      rows.push([d+1,staff.id,staff.name,kind,bw ? '出勤' : '休み',aw ? '出勤' : '休み',bu,au,actionName]);
      if (rows.length >= 195) break;
    }
    if (rows.length >= 195) break;
  }
  if (!rows.length) rows.push(['－','－','－','変更なし','','','','',actionName]);
  sheet.getRange(5,1,rows.length,9).setValues(rows);
}

/* =========================================================
 * 20. バックアップ・復元
 * ========================================================= */

function backupCurrentState_(actionName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.AUTO);
  const backupSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.BACKUP);

  backupSheet.clearContents();

  backupSheet.getRange('A1:B3').setValues([
    ['最新バックアップ', new Date()],
    ['処理', actionName],
    ['GASバージョン', SHIFT_APP.VERSION]
  ]);

  // 数式も含めてそのまま保存する。
  autoSheet
    .getRange(SHIFT_APP.RANGE.WORK_PLAN)
    .copyTo(
      backupSheet.getRange('A5:AE19'),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  autoSheet
    .getRange(SHIFT_APP.RANGE.UNIT_PLAN)
    .copyTo(
      backupSheet.getRange('A23:AE37'),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  appendLog_(
    'INFO',
    'バックアップ',
    actionName,
    '自動計算の勤務・単位グリッドを保存しました。'
  );
}


function restoreBackup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.AUTO);
  const backupSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.BACKUP);

  const marker = String(backupSheet.getRange('A1').getValue() || '').trim();

  if (marker !== '最新バックアップ') {
    throw new Error(
      '復元エラー：直前バックアップが見つかりません。'
    );
  }

  backupSheet
    .getRange('A5:AE19')
    .copyTo(
      autoSheet.getRange(SHIFT_APP.RANGE.WORK_PLAN),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  backupSheet
    .getRange('A23:AE37')
    .copyTo(
      autoSheet.getRange(SHIFT_APP.RANGE.UNIT_PLAN),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  SpreadsheetApp.flush();

  appendLog_(
    'INFO',
    'バックアップ復元',
    '',
    '直前状態へ復元しました。'
  );

  showAlert_(
    'バックアップ復元',
    '直前のシフト状態へ戻しました。'
  );
}


/* =========================================================
 * 21. 当日欠勤ロック
 * ========================================================= */

function addAbsenceLock_(ctx, day, staffId) {
  const lockSheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.LOCK);
  const range = lockSheet.getRange('A5:H80');
  const values = range.getValues();

  let targetRow = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === '') {
      targetRow = i + 5;
      break;
    }
  }

  if (targetRow < 0) {
    throw new Error(
      '入力エラー：「配置ロック」に空き行がありません。不要なロックを整理してください。'
    );
  }

  const date = new Date(
    ctx.targetMonth.getFullYear(),
    ctx.targetMonth.getMonth(),
    day
  );

  // D列「氏名」は数式列なので上書きしない。
  lockSheet.getRange(targetRow, 1, 1, 3).setValues([[
    '○',
    date,
    staffId
  ]]);

  lockSheet.getRange(targetRow, 5, 1, 3).setValues([[
    '休み固定',
    '',
    '当日欠勤'
  ]]);

  appendLog_(
    'WARN',
    '当日欠勤',
    staffId + ' / ' + day + '日',
    '配置ロックへ休み固定を追加しました。'
  );
}


/* =========================================================
 * 22. 現在プラン読込
 * ========================================================= */

function readCurrentPlan_(ctx) {
  const autoSheet = getSheetOrThrow_(ctx.ss, SHIFT_APP.SHEETS.AUTO);

  const rawWork = autoSheet
    .getRange(SHIFT_APP.RANGE.WORK_PLAN)
    .getValues();

  const rawUnit = autoSheet
    .getRange(SHIFT_APP.RANGE.UNIT_PLAN)
    .getDisplayValues();

  const work = matrix_(
    SHIFT_APP.MAX_STAFF,
    SHIFT_APP.MAX_DAYS,
    0
  );

  const unit = matrix_(
    SHIFT_APP.MAX_STAFF,
    SHIFT_APP.MAX_DAYS,
    ''
  );

  for (let s = 0; s < SHIFT_APP.MAX_STAFF; s++) {
    for (let d = 0; d < SHIFT_APP.MAX_DAYS; d++) {
      work[s][d] = Number(rawWork[s][d]) === 1 ? 1 : 0;
      unit[s][d] = normalizeUnit_(rawUnit[s][d]);
    }
  }

  return {
    work: work,
    unit: unit
  };
}


/* =========================================================
 * 23. 配置設定の解析
 * ========================================================= */

function parsePlacementSettings_(values, unitCount) {
  const result = {
    '第1': createEmptyUnitConfig_(),
    '第2': createEmptyUnitConfig_()
  };

  let currentUnit = '';

  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const a = String(row[0] || '').trim();

    if (a.indexOf('第1教室') >= 0) {
      currentUnit = '第1';
      continue;
    }

    if (a.indexOf('第2教室') >= 0) {
      currentUnit = '第2';
      continue;
    }

    if (!currentUnit) continue;

    // 次の大項目へ移動したら単位セクション終了。
    if (/^\d+\./.test(a) &&
        a.indexOf('第1教室') < 0 &&
        a.indexOf('第2教室') < 0) {
      currentUnit = '';
      continue;
    }

    if (!result[currentUnit]) continue;

    const entry = {
      main: String(row[1] || '').trim(),
      sub: String(row[2] || '').trim(),
      operation: String(row[3] || '').trim(),
      placementType: String(row[4] || '').trim(),
      experience: String(row[5] || '').trim()
    };

    if (a === '児発管') {
      result[currentUnit].jihatsukan = entry;
    } else if (a === '児童指導員等加配') {
      result[currentUnit].addOn = entry;
    } else if (a === '専門的支援実施') {
      result[currentUnit].specialExecution = entry;
    } else if (a === '専門的支援体制加算') {
      result[currentUnit].specialSystem = entry;
    } else if (a === '最低人員優先①') {
      result[currentUnit].minimumPriority1 = entry;
    } else if (a === '最低人員優先②') {
      result[currentUnit].minimumPriority2 = entry;
    }
  }

  if (unitCount < 2) {
    result['第2'] = createEmptyUnitConfig_();
  }

  return result;
}


function createEmptyUnitConfig_() {
  return {
    jihatsukan: null,
    addOn: null,
    specialExecution: null,
    specialSystem: null,
    minimumPriority1: null,
    minimumPriority2: null
  };
}


/* =========================================================
 * 24. 単位・候補補助
 * ========================================================= */

function getAssignedUnit_(ctx, state, staffIndex, dayIndex) {
  const override = normalizeUnit_(state.unit[staffIndex][dayIndex]);

  if (override === '第1' || override === '第2') {
    return override;
  }

  const home = normalizeUnit_(ctx.staff[staffIndex].unit);

  if (home === '第1' || home === '第2') {
    return home;
  }

  return '';
}


function getAssignedUnitFromMatrices_(
  staff,
  unitMatrix,
  staffIndex,
  dayIndex
) {
  const override = normalizeUnit_(
    unitMatrix[staffIndex][dayIndex]
  );

  if (override === '第1' || override === '第2') {
    return override;
  }

  const home = normalizeUnit_(staff.unit);

  if (home === '第1' || home === '第2') {
    return home;
  }

  return '';
}


function setAssignedUnit_(ctx, state, staffIndex, dayIndex, unit) {
  const normalized = normalizeUnit_(unit);

  if (normalized !== '第1' && normalized !== '第2') {
    return;
  }

  state.unit[staffIndex][dayIndex] = normalized;
}


function chooseUnitByDeficit_(ctx, state, dayIndex) {
  const need1 = Math.max(
    0,
    requiredDirectSupportCount_(ctx, ctx.days[dayIndex], '第1') -
    countLegalStaff_(ctx, state, dayIndex, '第1')
  );

  if (ctx.unitCount < 2) {
    return '第1';
  }

  const need2 = Math.max(
    0,
    requiredDirectSupportCount_(ctx, ctx.days[dayIndex], '第2') -
    countLegalStaff_(ctx, state, dayIndex, '第2')
  );

  return need1 >= need2 ? '第1' : '第2';
}


/* =========================================================
 * 25. 連勤計算補助
 * ========================================================= */

function wouldExceedMaxConsecutive_(workRow, dayIndex, maxConsecutive) {
  if (workRow[dayIndex] === 1) {
    return getMaxConsecutive_(workRow, SHIFT_APP.MAX_DAYS) > maxConsecutive;
  }

  const test = workRow.slice();
  test[dayIndex] = 1;

  return getMaxConsecutive_(
    test,
    SHIFT_APP.MAX_DAYS
  ) > maxConsecutive;
}


function getMaxConsecutive_(workRow, daysInMonth) {
  let max = 0;
  let current = 0;

  for (let d = 0; d < daysInMonth; d++) {
    if (Number(workRow[d]) === 1) {
      current++;
      if (current > max) max = current;
    } else {
      current = 0;
    }
  }

  return max;
}


function findFirstExcessRun_(workRow, daysInMonth, maxConsecutive) {
  let start = -1;
  let current = 0;

  for (let d = 0; d < daysInMonth; d++) {
    if (Number(workRow[d]) === 1) {
      if (current === 0) start = d;
      current++;

      if (current > maxConsecutive) {
        let end = d;

        while (end + 1 < daysInMonth &&
               Number(workRow[end + 1]) === 1) {
          end++;
        }

        return {
          start: start,
          end: end
        };
      }
    } else {
      current = 0;
      start = -1;
    }
  }

  return null;
}


/* =========================================================
 * 26. ログ
 * ========================================================= */

function appendLog_(level, action, target, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.LOG);

  ensureLogHeader_(logSheet);

  logSheet.appendRow([
    new Date(),
    level,
    action,
    target || '',
    message || '',
    SHIFT_APP.VERSION
  ]);
}


function ensureLogHeader_(sheet) {
  const current = sheet
    .getRange(1, 1, 1, SHIFT_APP.LOG_HEADERS.length)
    .getDisplayValues()[0];

  const isEmpty = current.every(function (x) {
    return String(x || '').trim() === '';
  });

  if (isEmpty) {
    sheet
      .getRange(1, 1, 1, SHIFT_APP.LOG_HEADERS.length)
      .setValues([SHIFT_APP.LOG_HEADERS]);
  }
}


function writeValidationLog_(action, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getSheetOrThrow_(ss, SHIFT_APP.SHEETS.LOG);
  ensureLogHeader_(logSheet);

  const rows = [];
  const version = SHIFT_APP.VERSION;

  if (result.errors.length === 0 &&
      result.warnings.length === 0) {
    rows.push([
      new Date(), 'INFO', action, '', 'エラー・警告なし', version
    ]);
  } else {
    result.errors.forEach(function (message) {
      rows.push([new Date(), 'ERROR', action, '', message, version]);
    });

    result.warnings.forEach(function (message) {
      rows.push([new Date(), 'WARN', action, '', message, version]);
    });
  }

  if (rows.length > 0) {
    logSheet
      .getRange(logSheet.getLastRow() + 1, 1, rows.length, 6)
      .setValues(rows);
  }
}


/* =========================================================
 * 27. UI・エラー処理
 * ========================================================= */

function runSafely_(actionName, callback) {
  try {
    callback();
  } catch (error) {
    const message = error && error.message
      ? error.message
      : String(error);

    try {
      appendLog_(
        'ERROR',
        actionName,
        '',
        message
      );
    } catch (logError) {
      // ログシート自体が壊れている場合でも、本来のエラー表示を優先する。
    }

    showAlert_(
      'エラー',
      message
    );

    throw error;
  }
}


function showAlert_(title, message) {
  SpreadsheetApp.getUi().alert(
    title,
    String(message || ''),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function formatValidationResult_(result) {
  const lines = [];

  if (result.errors.length === 0) {
    lines.push('【重大エラー】0件');
  } else {
    lines.push('【重大エラー】' + result.errors.length + '件');
    result.errors.slice(0, 20).forEach(function (x) {
      lines.push('・' + x);
    });

    if (result.errors.length > 20) {
      lines.push('・ほか ' + (result.errors.length - 20) + '件');
    }
  }

  lines.push('');

  if (result.warnings.length === 0) {
    lines.push('【警告】0件');
  } else {
    lines.push('【警告】' + result.warnings.length + '件');
    result.warnings.slice(0, 20).forEach(function (x) {
      lines.push('・' + x);
    });

    if (result.warnings.length > 20) {
      lines.push('・ほか ' + (result.warnings.length - 20) + '件');
    }
  }

  return lines.join('\n');
}


/* =========================================================
 * 28. 共通ユーティリティ
 * ========================================================= */

function getSheetOrThrow_(ss, name) {
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    throw new Error(
      '構成エラー：必要なシート「' + name + '」が見つかりません。'
    );
  }

  return sheet;
}


function buildHeaderMap_(headers) {
  const map = {};

  headers.forEach(function (header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });

  return map;
}


function valueAt_(row, headerMap, headerName) {
  if (!Object.prototype.hasOwnProperty.call(headerMap, headerName)) {
    return '';
  }

  return row[headerMap[headerName]];
}


function stringAt_(row, headerMap, headerName) {
  return String(
    valueAt_(row, headerMap, headerName) || ''
  ).trim();
}


function numberAt_(row, headerMap, headerName) {
  return numberOrZero_(
    valueAt_(row, headerMap, headerName)
  );
}


function numberOrZero_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return 0;
  }

  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}


function isYes_(value) {
  const text = String(value || '').trim();

  return text === '○' ||
         text === '可' ||
         text === 'はい' ||
         text === 'TRUE' ||
         text === 'true';
}


function normalizeUnit_(value) {
  const text = String(value || '').trim();

  if (text === '第1' ||
      text === '第1教室' ||
      text === '1' ||
      text === '1単位') {
    return '第1';
  }

  if (text === '第2' ||
      text === '第2教室' ||
      text === '2' ||
      text === '2単位') {
    return '第2';
  }

  if (text === '共通') {
    return '共通';
  }

  return '';
}


function firstDayOfMonth_(value) {
  const date = normalizeDate_(value);

  if (!date) return null;

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    1
  );
}


function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    // Google Sheetsの日付シリアル値（例: 46296）にも対応する。
    // ミリ秒タイムスタンプと誤認すると1970年になるため、
    // 1000000未満の正数は Sheets/Excel 系の日付シリアルとして扱う。
    if (value > 0 && value < 1000000) {
      const wholeDays = Math.floor(value);
      const utcMillis = Date.UTC(1899, 11, 30) + wholeDays * 86400000;
      const serialDate = new Date(utcMillis);

      if (!isNaN(serialDate.getTime())) {
        return new Date(
          serialDate.getUTCFullYear(),
          serialDate.getUTCMonth(),
          serialDate.getUTCDate()
        );
      }
    }

    const date = new Date(value);
    if (!isNaN(date.getTime())) return date;
  }

  const text = String(value || '').trim();
  if (!text) return null;

  // 「2026年10月」「2026/10」「2026-10」も安全に解釈する。
  const ym = text.match(/^(\d{4})\s*(?:年|[\/-])\s*(\d{1,2})\s*(?:月)?$/);
  if (ym) {
    const year = Number(ym[1]);
    const month = Number(ym[2]);
    if (month >= 1 && month <= 12) {
      return new Date(year, month - 1, 1);
    }
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
}


function findLabelValue_(values, label) {
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length - 1; c++) {
      if (String(values[r][c] || '').trim() === label) {
        return values[r][c + 1];
      }
    }
  }

  return '';
}


function findStaffByRef_(ctx, ref) {
  const text = String(ref || '').trim();
  if (!text) return null;

  return ctx.staff.find(function (staff) {
    return staff.active &&
      (staff.id === text || staff.name === text);
  }) || null;
}


function weekdayJa_(date) {
  return ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
}


function parseWeekdays_(text) {
  const value = String(text || '').trim();

  if (!value) {
    return [];
  }

  // よく使う表現を先に処理する。
  if (value === '平日' ||
      value.indexOf('平日のみ') >= 0 ||
      value.indexOf('月～金') >= 0 ||
      value.indexOf('月-金') >= 0) {
    return ['月', '火', '水', '木', '金'];
  }

  if (value === '土日' ||
      value.indexOf('土日のみ') >= 0) {
    return ['土', '日'];
  }

  const days = ['月', '火', '水', '木', '金', '土', '日'];

  return days.filter(function (day) {
    return value.indexOf(day) >= 0;
  });
}


function matrix_(rows, cols, initialValue) {
  const result = [];

  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      row.push(initialValue);
    }
    result.push(row);
  }

  return result;
}


function countWorkDays_(workRow, daysInMonth) {
  let count = 0;

  for (let d = 0; d < daysInMonth; d++) {
    if (Number(workRow[d]) === 1) count++;
  }

  return count;
}


function unique_(array) {
  return Array.from(new Set(array));
}