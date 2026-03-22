/* ==========================================
   レシートぱっと - メインアプリケーション
   レシート画像 → OCR → 勘定科目自動判定 → CSV出力
   ========================================== */

(() => {
  'use strict';

  // =========================================
  // 勘定科目 自動判定ルール
  // 店名のキーワードに基づいて勘定科目を推定する
  // =========================================
  const CATEGORY_RULES = [
    // 旅費交通費
    { keywords: ['タクシー', 'taxi', 'JR', '鉄道', '電車', 'バス', 'Suica', 'PASMO', 'ICOCA', '交通', '高速', '駐車場', 'パーキング', 'ガソリン', '出光', 'ENEOS', 'エネオス', 'コスモ', 'シェル', '宇佐美'], category: '旅費交通費' },

    // 通信費
    { keywords: ['NTT', 'ドコモ', 'docomo', 'au', 'ソフトバンク', 'SoftBank', '楽天モバイル', 'UQ', 'Y!mobile', 'ワイモバイル', 'povo', 'LINEMO', 'ahamo', 'プロバイダ', 'インターネット', 'Wi-Fi', 'さくら', 'AWS', 'サーバー', 'レンタルサーバ', 'ムームー', 'お名前'], category: '通信費' },

    // 接待交際費
    { keywords: ['居酒屋', '焼肉', '焼鳥', 'バー', 'BAR', '寿司', 'すし', '鮨', '和食', 'イタリアン', 'フレンチ', '中華', 'レストラン', '料亭', 'ダイニング', '宴会', 'ホテル'], category: '接待交際費' },

    // 会議費
    { keywords: ['スターバックス', 'スタバ', 'Starbucks', 'タリーズ', 'Tully', 'ドトール', 'DOUTOR', 'コメダ', 'カフェ', 'CAFE', 'café', 'サンマルク', 'ベローチェ', 'プロント', 'PRONTO', '上島珈琲', '珈琲館', 'ブルーボトル'], category: '会議費' },

    // 新聞図書費
    { keywords: ['書店', '本屋', 'Amazon', 'アマゾン', '紀伊國屋', '丸善', 'ジュンク堂', 'TSUTAYA', 'ブックオフ', 'Kindle', '日経', '新聞', '朝日', '読売', '毎日', 'Udemy', 'Coursera'], category: '新聞図書費' },

    // 水道光熱費
    { keywords: ['電力', '東京電力', '関西電力', '中部電力', 'ガス', '東京ガス', '大阪ガス', '水道', '上下水道'], category: '水道光熱費' },

    // 荷造運賃
    { keywords: ['ヤマト', '佐川', '日本郵便', '郵便局', 'ゆうパック', 'クロネコ', 'FedEx', 'DHL', '切手', 'レターパック'], category: '荷造運賃' },

    // 広告宣伝費
    { keywords: ['Google Ads', 'Facebook', 'Instagram', 'Twitter', '広告', 'チラシ', '印刷', 'ラクスル', 'プリントパック', '名刺'], category: '広告宣伝費' },

    // 消耗品費（幅広く）
    { keywords: ['コンビニ', 'セブン', 'ファミリーマート', 'ファミマ', 'ローソン', 'ミニストップ', 'デイリー', '百均', '100均', 'ダイソー', 'セリア', 'キャンドゥ', 'ホームセンター', 'カインズ', 'コーナン', 'ヨドバシ', 'ビックカメラ', 'ヤマダ電機', 'エディオン', 'ケーズ', 'Amazon', '楽天', 'ドラッグ', 'マツキヨ', 'ウエルシア', '文具', 'コクヨ', 'ロフト', 'LOFT', '東急ハンズ', 'ハンズ', 'ニトリ', 'IKEA', 'イケア', '無印', 'MUJI', 'ドン・キホーテ', 'ドンキ'], category: '消耗品費' },
  ];

  // =========================================
  // レシートデータ管理
  // =========================================
  let receipts = [];

  // ローカルストレージから復元
  function loadReceipts() {
    try {
      const saved = localStorage.getItem('receipts_data');
      if (saved) {
        receipts = JSON.parse(saved);
      }
    } catch (e) {
      receipts = [];
    }
  }

  function saveReceipts() {
    try {
      localStorage.setItem('receipts_data', JSON.stringify(receipts));
    } catch (e) {
      // ストレージフルの場合は無視
    }
  }

  // =========================================
  // DOM要素の取得
  // =========================================
  const uploadZone = document.getElementById('uploadZone');
  const fileInput = document.getElementById('fileInput');
  const cameraInput = document.getElementById('cameraInput');
  const processingSection = document.getElementById('processingSection');
  const processingTitle = document.getElementById('processingTitle');
  const processingDetail = document.getElementById('processingDetail');
  const processingBarFill = document.getElementById('processingBarFill');
  const editSection = document.getElementById('editSection');
  const editPreviewImg = document.getElementById('editPreviewImg');
  const editDate = document.getElementById('editDate');
  const editAmount = document.getElementById('editAmount');
  const editStore = document.getElementById('editStore');
  const editCategory = document.getElementById('editCategory');
  const editMemo = document.getElementById('editMemo');
  const editSave = document.getElementById('editSave');
  const editCancel = document.getElementById('editCancel');
  const listSection = document.getElementById('listSection');
  const listSummary = document.getElementById('listSummary');
  const listItems = document.getElementById('listItems');
  const receiptCount = document.getElementById('receiptCount');
  const exportBtn = document.getElementById('exportBtn');
  const exportFormat = document.getElementById('exportFormat');

  // =========================================
  // ドラッグ&ドロップ対応
  // =========================================
  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  });

  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('dragover');
  });

  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (files.length > 0) {
      processFiles(files);
    }
  });

  // ファイル選択
  fileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      processFiles(files);
      fileInput.value = ''; // リセット
    }
  });

  // カメラ撮影
  cameraInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      processFiles(files);
      cameraInput.value = '';
    }
  });

  // =========================================
  // 画像ファイルの処理
  // =========================================
  let processingQueue = [];
  let currentImageDataUrl = '';

  async function processFiles(files) {
    processingQueue = [...processingQueue, ...files];
    if (processingQueue.length === files.length) {
      processNext();
    }
  }

  async function processNext() {
    if (processingQueue.length === 0) return;

    const file = processingQueue.shift();

    // 画像をDataURLに変換
    const dataUrl = await fileToDataUrl(file);
    currentImageDataUrl = dataUrl;

    // OCR処理の開始
    showProcessing();

    try {
      const ocrResult = await runOCR(dataUrl);
      const parsed = parseOCRText(ocrResult);

      // 編集フォームに結果を表示
      showEditForm(dataUrl, parsed);

    } catch (error) {
      hideProcessing();
      // OCRが失敗しても手動入力できるように編集フォームを表示
      showEditForm(dataUrl, {
        date: new Date().toISOString().split('T')[0],
        amount: '',
        store: '',
        category: '消耗品費'
      });
    }
  }

  function fileToDataUrl(file) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.readAsDataURL(file);
    });
  }

  // =========================================
  // OCR処理（Tesseract.js）
  // =========================================
  async function runOCR(imageDataUrl) {
    processingTitle.textContent = 'レシートを読み取り中...';
    processingDetail.textContent = 'OCRエンジンを準備しています';
    processingBarFill.style.width = '10%';

    try {
      const result = await Tesseract.recognize(
        imageDataUrl,
        'jpn+eng',
        {
          logger: (info) => {
            if (info.status === 'recognizing text') {
              const progress = Math.round(info.progress * 100);
              processingBarFill.style.width = `${10 + progress * 0.85}%`;
              processingDetail.textContent = `文字認識中... ${progress}%`;
            } else if (info.status === 'loading language traineddata') {
              processingDetail.textContent = '日本語認識データを読込中...';
              processingBarFill.style.width = '5%';
            }
          }
        }
      );

      processingBarFill.style.width = '100%';
      processingDetail.textContent = '解析完了！';

      return result.data.text;
    } catch (error) {
      throw new Error('OCR処理に失敗しました: ' + error.message);
    }
  }

  // =========================================
  // OCR結果からデータを抽出する
  // =========================================
  function parseOCRText(text) {
    const result = {
      date: '',
      amount: '',
      store: '',
      category: '消耗品費'
    };

    if (!text || text.trim().length === 0) {
      result.date = new Date().toISOString().split('T')[0];
      return result;
    }

    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

    // --- 日付の抽出 ---
    // 様々な日付フォーマットに対応
    const datePatterns = [
      /(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/,                // 2024/03/15, 2024年3月15日
      /([Rr令]?\d{1,2})[\/\-\.年](\d{1,2})[\/\-\.月](\d{1,2})/,   // R6/03/15, 令6年3月15日
      /(\d{1,2})[\/\-月](\d{1,2})[\/\-日]/                         // 3/15, 3月15日
    ];

    for (const line of lines) {
      for (const pattern of datePatterns) {
        const match = line.match(pattern);
        if (match) {
          let year, month, day;

          if (pattern === datePatterns[0]) {
            year = parseInt(match[1]);
            month = parseInt(match[2]);
            day = parseInt(match[3]);
          } else if (pattern === datePatterns[1]) {
            // 令和変換
            const reiwa = parseInt(match[1].replace(/[Rr令]/g, ''));
            year = reiwa > 100 ? reiwa : 2018 + reiwa;
            month = parseInt(match[2]);
            day = parseInt(match[3]);
          } else {
            year = new Date().getFullYear();
            month = parseInt(match[1]);
            day = parseInt(match[2]);
          }

          if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            result.date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            break;
          }
        }
      }
      if (result.date) break;
    }

    if (!result.date) {
      result.date = new Date().toISOString().split('T')[0];
    }

    // --- 金額の抽出 ---
    // 「合計」「計」「お買上」などのキーワード付き金額を優先
    const amountPatterns = [
      /(?:合計|計|お買[い上]?げ?|お支払[い]?|税込|総額|請求)\s*[¥￥]?\s*([0-9,]+)/,
      /[¥￥]\s*([0-9,]{3,})/,
      /([0-9,]{3,})\s*円/,
    ];

    let bestAmount = '';
    let bestAmountPriority = -1;

    for (const line of lines) {
      for (let i = 0; i < amountPatterns.length; i++) {
        const match = line.match(amountPatterns[i]);
        if (match) {
          const amount = match[1].replace(/,/g, '');
          const amountNum = parseInt(amount);
          // より優先度の高いパターンか、同じ優先度でより大きい金額を採用
          if (i < bestAmountPriority || bestAmountPriority === -1 ||
              (i === bestAmountPriority && amountNum > parseInt(bestAmount || '0'))) {
            bestAmount = amount;
            bestAmountPriority = i;
          }
        }
      }
    }

    result.amount = bestAmount;

    // --- 店名の抽出 ---
    // 通常レシートの最初の数行に店名がある
    const storeExcludes = ['レシート', '領収', '精算', '明細', '伝票', '電話', 'tel', 'TEL', 'http', 'www'];
    for (const line of lines.slice(0, 5)) {
      const cleaned = line.replace(/[\s　]+/g, '').trim();
      if (cleaned.length < 2 || cleaned.length > 30) continue;
      if (/^\d+$/.test(cleaned)) continue;
      if (storeExcludes.some(ex => cleaned.toLowerCase().includes(ex.toLowerCase()))) continue;
      if (/^[0-9\-\/\.:]+$/.test(cleaned)) continue;

      result.store = cleaned;
      break;
    }

    // --- 勘定科目の判定 ---
    result.category = guessCategory(result.store, text);

    return result;
  }

  // =========================================
  // 勘定科目の自動判定
  // =========================================
  function guessCategory(storeName, fullText) {
    const searchText = (storeName + ' ' + fullText).toLowerCase();

    for (const rule of CATEGORY_RULES) {
      for (const keyword of rule.keywords) {
        if (searchText.includes(keyword.toLowerCase())) {
          return rule.category;
        }
      }
    }

    return '消耗品費'; // デフォルト
  }

  // =========================================
  // UI表示制御
  // =========================================
  function showProcessing() {
    processingSection.style.display = 'block';
    editSection.style.display = 'none';
    processingBarFill.style.width = '0%';
  }

  function hideProcessing() {
    processingSection.style.display = 'none';
  }

  function showEditForm(imageDataUrl, data) {
    hideProcessing();
    editSection.style.display = 'block';
    editPreviewImg.src = imageDataUrl;
    editDate.value = data.date || '';
    editAmount.value = data.amount || '';
    editStore.value = data.store || '';
    editCategory.value = data.category || '消耗品費';
    editMemo.value = '';

    // フォーカスを金額欄に（OCR結果が不正確な場合が多いため）
    setTimeout(() => editAmount.focus(), 200);
  }

  function hideEditForm() {
    editSection.style.display = 'none';
    currentImageDataUrl = '';
  }

  // 店名変更時に勘定科目を再判定
  editStore.addEventListener('input', () => {
    const guessed = guessCategory(editStore.value, '');
    editCategory.value = guessed;
  });

  // =========================================
  // レシート登録
  // =========================================
  editSave.addEventListener('click', () => {
    const date = editDate.value;
    const amount = editAmount.value;
    const store = editStore.value.trim();

    if (!date || !amount || !store) {
      // 未入力項目のハイライト
      if (!date) editDate.style.borderColor = '#F87171';
      if (!amount) editAmount.style.borderColor = '#F87171';
      if (!store) editStore.style.borderColor = '#F87171';
      setTimeout(() => {
        editDate.style.borderColor = '';
        editAmount.style.borderColor = '';
        editStore.style.borderColor = '';
      }, 2000);
      return;
    }

    const receipt = {
      id: Date.now(),
      date: date,
      amount: parseInt(amount),
      store: store,
      category: editCategory.value,
      memo: editMemo.value.trim()
    };

    receipts.push(receipt);
    saveReceipts();

    hideEditForm();
    renderList();

    // 次のファイルがあれば処理
    if (processingQueue.length > 0) {
      processNext();
    }
  });

  editCancel.addEventListener('click', () => {
    hideEditForm();
    processingQueue = []; // キューもクリア
  });

  // =========================================
  // レシート一覧表示
  // =========================================
  function renderList() {
    if (receipts.length === 0) {
      listSection.style.display = 'none';
      receiptCount.textContent = '0件';
      exportBtn.disabled = true;
      return;
    }

    listSection.style.display = 'block';
    receiptCount.textContent = `${receipts.length}件`;
    exportBtn.disabled = false;

    // サマリー
    const totalAmount = receipts.reduce((sum, r) => sum + r.amount, 0);
    listSummary.innerHTML = `
      <div class="list-summary-item">
        登録数 <strong>${receipts.length}</strong> 件
      </div>
      <div class="list-summary-item">
        合計金額 <strong>¥${totalAmount.toLocaleString()}</strong>
      </div>
    `;

    // 日付の新しい順にソート
    const sorted = [...receipts].sort((a, b) => b.date.localeCompare(a.date));

    listItems.innerHTML = '';
    for (const r of sorted) {
      const item = document.createElement('div');
      item.className = 'receipt-item';

      // 日付を表示用にフォーマット
      const dateParts = r.date.split('-');
      const displayDate = `${dateParts[0]}/${dateParts[1]}/${dateParts[2]}`;

      item.innerHTML = `
        <div class="receipt-category-badge">${escapeHtml(r.category)}</div>
        <div class="receipt-detail">
          <div class="receipt-store">${escapeHtml(r.store)}${r.memo ? ' <span style="color:#6B7280;font-size:11px">— ' + escapeHtml(r.memo) + '</span>' : ''}</div>
          <div class="receipt-date">${displayDate}</div>
        </div>
        <div class="receipt-amount">¥${r.amount.toLocaleString()}</div>
        <button class="receipt-delete" data-id="${r.id}" title="削除">🗑</button>
      `;

      listItems.appendChild(item);
    }

    // 削除ボタンのイベント
    listItems.querySelectorAll('.receipt-delete').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = parseInt(btn.dataset.id);
        if (confirm('このレシートを削除しますか？')) {
          receipts = receipts.filter(r => r.id !== id);
          saveReceipts();
          renderList();
        }
      });
    });
  }

  // =========================================
  // CSV出力
  // =========================================
  exportBtn.addEventListener('click', () => {
    if (receipts.length === 0) return;

    const format = exportFormat.value;
    let csv = '';

    switch (format) {
      case 'freee':
        csv = generateFreeeCSV();
        break;
      case 'yayoi':
        csv = generateYayoiCSV();
        break;
      case 'moneyforward':
        csv = generateMoneyForwardCSV();
        break;
      default:
        csv = generateGenericCSV();
    }

    downloadCSV(csv, format);
  });

  /**
   * freee形式のCSVを生成
   * freeeの「自動仕訳」インポート形式
   */
  function generateFreeeCSV() {
    const header = '収支区分,管理番号,発生日,決済期日,取引先,勘定科目,税区分,金額,備考\n';
    const rows = receipts.map(r => {
      return [
        '支出',           // 収支区分
        '',               // 管理番号
        r.date,           // 発生日
        r.date,           // 決済期日
        r.store,          // 取引先
        r.category,       // 勘定科目
        '課税仕入10%',     // 税区分
        r.amount,         // 金額
        r.memo            // 備考
      ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    }).join('\n');

    return header + rows;
  }

  /**
   * 弥生形式のCSVを生成
   * 弥生の仕訳インポート形式
   */
  function generateYayoiCSV() {
    const header = '日付,借方勘定科目,借方金額,貸方勘定科目,貸方金額,摘要\n';
    const rows = receipts.map(r => {
      return [
        r.date,           // 日付
        r.category,       // 借方勘定科目
        r.amount,         // 借方金額
        '現金',           // 貸方勘定科目
        r.amount,         // 貸方金額
        `${r.store}${r.memo ? ' ' + r.memo : ''}`  // 摘要
      ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    }).join('\n');

    return header + rows;
  }

  /**
   * マネーフォワード形式のCSVを生成
   */
  function generateMoneyForwardCSV() {
    const header = '日付,内容,金額,大項目,中項目,メモ\n';
    const rows = receipts.map(r => {
      return [
        r.date,
        r.store,
        r.amount,
        r.category,
        '',
        r.memo
      ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    }).join('\n');

    return header + rows;
  }

  /**
   * 汎用CSV
   */
  function generateGenericCSV() {
    const header = '日付,店名,金額,勘定科目,メモ\n';
    const rows = receipts.map(r => {
      return [r.date, r.store, r.amount, r.category, r.memo]
        .map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    }).join('\n');

    return header + rows;
  }

  /**
   * CSVファイルをダウンロードする
   */
  function downloadCSV(csvContent, format) {
    const formatNames = {
      freee: 'freee',
      yayoi: '弥生',
      moneyforward: 'マネーフォワード',
      generic: '汎用'
    };

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const filename = `レシートぱっと_${formatNames[format]}_${dateStr}.csv`;

    // BOM付きUTF-8でダウンロード
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // =========================================
  // ユーティリティ
  // =========================================
  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // =========================================
  // 初期化
  // =========================================
  loadReceipts();
  renderList();

})();
