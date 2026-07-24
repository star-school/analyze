// 外部ライブラリ (xlsx, Chart.js) がグローバルに読み込まれている前提

let rawData = []; // Excelから読み込んだ生のJSON行データ
let processedData = []; // 列マッピングおよび学年補正済みのデータ
let filteredData = []; // フィルター適用後のデータ
let allStoreFilteredData = []; // 教室フィルタを無視したフィルタ適用後のデータ（KPIマトリクス用）
let charts = {}; // 作成されたグラフオブジェクトの参照
let storeTableSort = { key: 'total', direction: 'desc' }; // 教室別テーブルのソート状態
let detailTableSort = { key: 'id', direction: 'asc' }; // 生徒データ詳細テーブルのソート状態
let yoyMode = 'inq'; // 前年同期比グラフの表示モード ('inq': 問合せ数, 'enroll': 入会数)
let currentStoreSummaryList = []; // Excel出力用にマトリクス集計結果を保持するグローバル変数

// 列の同義語辞書（表記揺れ対応用）
const COLUMN_SYNONYMS = {
    status: ['区分', 'ステータス', '対応区分', '結果'],
    store: ['校舎名', '校舎', '教室名', '教室', '教場名', '教場ID'],
    currentGrade: ['学年', '現在の学年', '学年（現在）'],
    purpose: ['来校目的', '問合せ区分', '目的', '問い合わせ区分', '受講コース', 'コース名'],
    inquiryDate: ['問合せ日', '問合せ日時', '問い合わせ日', '受付日'],
    trialDate: ['体験日時', '体験日', '体験日付', '体験時間'],
    studentId: ['受講生ID', '生徒ID', '生徒コード', '受講生コード'],
    enrollDate: ['入会日', '入会日付', '入会年月日'],
    studyTime: ['先月末までの総受講時間', '総受講時間', '受講時間'],
    lastName: ['姓', '苗字'],
    firstName: ['名', '名前'],
    fullName: ['氏名', '生徒名', '生徒氏名'],
    courseType: ['受講種別', '受講区分'],
    generation: ['年代', '年齢区分', '年齢層', '世代']
};

document.addEventListener('DOMContentLoaded', () => {
    initFileLoader();
    initFilters();
    initStoreTableSort();
    initDetailTableSort();
    
    // 基準日の初期化（本日）と変更イベント
    const baseDateInput = document.getElementById('base-date');
    if (baseDateInput) {
        baseDateInput.value = new Date().toISOString().split('T')[0];
        baseDateInput.addEventListener('change', () => {
            if (processedData.length > 0) {
                // 学年を再計算
                processedData.forEach(item => {
                    const dateForGrade = item.trialDateStr || item.inquiryDateStr;
                    item.trialGrade = calculateTrialGrade(item.currentGrade, dateForGrade, item.courseType, item.generation);
                });
                // フィルターと表示の更新（プルダウンの選択状態は維持）
                applyFiltersAndRender();
            }
        });
    }

    // Excel出力ボタンのイベント
    document.getElementById('export-excel-btn').addEventListener('click', exportStoreTableToExcel);
    const exportDetailBtn = document.getElementById('export-detail-excel-btn');
    if (exportDetailBtn) {
        exportDetailBtn.addEventListener('click', exportDetailTableToExcel);
    }
    
    // フィルターセクションの表示/非表示切り替え
    const toggleFilterBtn = document.getElementById('toggle-filter-btn');
    const filterSection = document.getElementById('filter-section');
    if (toggleFilterBtn && filterSection) {
        toggleFilterBtn.addEventListener('click', () => {
            const isCollapsed = filterSection.classList.contains('collapsed');
            if (isCollapsed) {
                filterSection.classList.remove('collapsed');
                toggleFilterBtn.textContent = '🔍 フィルター非表示';
                toggleFilterBtn.classList.remove('collapsed');
                toggleFilterBtn.style.background = 'var(--secondary-hover)';
                toggleFilterBtn.style.color = 'white';
            } else {
                filterSection.classList.add('collapsed');
                toggleFilterBtn.textContent = '🔍 フィルター表示';
                toggleFilterBtn.classList.add('collapsed');
                toggleFilterBtn.style.background = 'var(--primary-light)';
                toggleFilterBtn.style.color = 'var(--primary-hover)';
            }
        });
    }
    
    // チャートデータ表のトグル＆エクスポートの初期化
    initChartTableToggles();

    // AIチャチャット機能のイベント
    document.getElementById('save-api-key-btn').addEventListener('click', saveApiKey);
    document.getElementById('chat-submit-btn').addEventListener('click', handleAiChatSubmit);
    document.getElementById('chat-input-text').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleAiChatSubmit();
    });
    
    // 保存済みのAPIキーがあれば復元
    const savedApiKey = localStorage.getItem('gemini_api_key');
    if (savedApiKey) {
        document.getElementById('gemini-api-key').value = savedApiKey;
        document.getElementById('api-key-status').textContent = '設定済み';
        document.getElementById('api-key-status').style.color = 'var(--success)';
    }
    
    // マニュアルモーダルの表示・非表示
    const manualModal = document.getElementById('manual-modal');
    if (manualModal) {
        document.getElementById('open-manual-btn').addEventListener('click', () => {
            manualModal.style.display = 'flex';
        });
        const openWelcomeBtn = document.getElementById('open-manual-btn-welcome');
        if (openWelcomeBtn) {
            openWelcomeBtn.addEventListener('click', () => {
                manualModal.style.display = 'flex';
            });
        }
        document.getElementById('close-manual-btn').addEventListener('click', () => {
            manualModal.style.display = 'none';
        });
        // モーダル外クリックで閉じる
        manualModal.addEventListener('click', (e) => {
            if (e.target === manualModal) manualModal.style.display = 'none';
        });
    }

    const pdfBtn = document.getElementById('export-pdf-btn');
    if (pdfBtn) {
        pdfBtn.addEventListener('click', () => {
            window.print();
        });
    }

    document.getElementById('change-file-btn').addEventListener('click', () => {
        showScreen('welcome-screen');
        // Reset state
        rawData = [];
        processedData = [];
        filteredData = [];
        storeTableSort = { key: 'total', direction: 'desc' }; // ソートの初期化
        destroyCharts();
    });
});

// 教室別比較テーブルのソートイベントリスナー登録
function initStoreTableSort() {
    const headers = document.querySelectorAll('#store-table th[data-sort]');
    headers.forEach(th => {
        th.addEventListener('click', () => {
            const key = th.getAttribute('data-sort');
            if (storeTableSort.key === key) {
                storeTableSort.direction = storeTableSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                storeTableSort.key = key;
                storeTableSort.direction = 'desc'; // 新規項目は降順デフォルト
            }
            
            // ソートインジケーター（矢印）の表示更新
            headers.forEach(h => {
                const icon = h.querySelector('.sort-icon');
                if (h.getAttribute('data-sort') === key) {
                    icon.textContent = storeTableSort.direction === 'asc' ? '▲' : '▼';
                } else {
                    icon.textContent = '';
                }
            });
            
            // テーブルのみ再描画
            renderStoreTable(allStoreFilteredData);
        });
    });
}

// 生徒データ詳細テーブルのソートイベントリスナー登録
function initDetailTableSort() {
    const headers = document.querySelectorAll('#detail-table th[data-sort-detail]');
    headers.forEach(th => {
        th.addEventListener('click', () => {
            const key = th.getAttribute('data-sort-detail');
            if (detailTableSort.key === key) {
                detailTableSort.direction = detailTableSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                detailTableSort.key = key;
                detailTableSort.direction = 'asc'; // 新規項目は昇順デフォルト
            }
            
            // ソートインジケーターの表示更新
            headers.forEach(h => {
                const icon = h.querySelector('.sort-icon');
                if (h.getAttribute('data-sort-detail') === key) {
                    icon.textContent = detailTableSort.direction === 'asc' ? '▲' : '▼';
                } else {
                    icon.textContent = '';
                }
            });
            
            // フィルター済みデータ全体を再ソートして描画
            renderDetailTable(filteredData);
        });
    });
}

// 画面表示の切り替え
function showScreen(screenId) {
    if (screenId === 'welcome-screen') {
        document.getElementById('welcome-screen').className = 'screenactive';
        document.getElementById('dashboard-screen').className = 'screendisable';
    } else {
        document.getElementById('welcome-screen').className = 'screendisable';
        document.getElementById('dashboard-screen').className = 'screenactive';
    }
}

// ==========================================================================
// 1. ファイル読み込み処理
// ==========================================================================
function initFileLoader() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    // ドラッグ＆ドロップイベント
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file && file.name.endsWith('.xlsx')) {
            handleFile(file);
        } else {
            alert('Excelファイル (.xlsx) をドラッグ＆ドロップしてください。');
        }
    });

    // ファイル選択ボタンイベント
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            handleFile(file);
        }
    });
}

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            
            // 最初のシートを選択
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            
            // 外部システム由来のExcelファイルで !ref (有効範囲) が破損している場合に対処するため、実際の範囲を再計算する
            if (sheet && sheet['!ref']) {
                let minC = 10000000, minR = 10000000, maxC = 0, maxR = 0;
                let hasData = false;
                for (const key of Object.keys(sheet)) {
                    if (key[0] === '!') continue;
                    const cell = XLSX.utils.decode_cell(key);
                    if (cell.c < minC) minC = cell.c;
                    if (cell.r < minR) minR = cell.r;
                    if (cell.c > maxC) maxC = cell.c;
                    if (cell.r > maxR) maxR = cell.r;
                    hasData = true;
                }
                if (hasData) {
                    sheet['!ref'] = XLSX.utils.encode_range({ s: { c: minC, r: minR }, e: { c: maxC, r: maxR } });
                }
            }

            // ヘッダーと行データをパース
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (rows.length < 2) {
                alert('Excelシートに十分なデータが含まれていません。');
                return;
            }

            const headers = rows[0];
            const dataRows = rows.slice(1);

            // ヘッダー名から必要な列の同義語マッピングを解決
            const resolvedMapping = resolveColumnMapping(headers);

            // 生データをパースして、統一オブジェクトに変換
            processedData = dataRows.map((row, idx) => {
                const item = { id: idx + 1 };
                
                // 動的列マッピングの適用
                Object.keys(resolvedMapping).forEach(key => {
                    const colIdx = resolvedMapping[key];
                    item[key] = colIdx !== -1 ? row[colIdx] : '';
                });

                // 値のフォーマット調整
                item.store = cleanStoreName(item.store);
                item.status = item.status ? item.status.toString().trim() : '(空白)';
                item.purpose = item.purpose ? item.purpose.toString().trim() : '(空白)';
                item.courseType = item.courseType ? item.courseType.toString().trim() : '(空白)';
                item.generation = item.generation ? item.generation.toString().trim() : '';
                
                // 名前データの結合
                let nameStr = '';
                if (item.fullName) {
                    nameStr = item.fullName.toString().trim();
                } else if (item.lastName || item.firstName) {
                    nameStr = `${item.lastName || ''} ${item.firstName || ''}`.trim();
                }
                item.displayName = nameStr || '(無記名)';
                
                // 日付のパース
                item.inquiryDateStr = formatDate(item.inquiryDate);
                item.trialDateStr = formatDate(item.trialDate);
                item.enrollDateStr = formatDate(item.enrollDate);

                // 体験当時の学年補正
                const dateForGrade = item.trialDateStr || item.inquiryDateStr;
                item.trialGrade = calculateTrialGrade(item.currentGrade, dateForGrade, item.courseType, item.generation);

                return item;
            }).filter(item => item.store || item.status || item.currentGrade); // 空白行の除外

            // ダッシュボード表示への切り替え
            showScreen('dashboard-screen');
            document.getElementById('dashboard-subtitle').innerText = `ファイル: ${file.name} (データ件数: ${processedData.length}件)`;

            // 問い合わせ日の最小値・最大値を算出して初期設定する
            const inquiryDates = processedData
                .map(item => item.inquiryDateStr)
                .filter(d => d && d !== '(空白)' && d !== '');
            
            if (inquiryDates.length > 0) {
                inquiryDates.sort(); // 日付文字列を昇順ソート
                document.getElementById('filter-date-start').value = inquiryDates[0];
                document.getElementById('filter-date-end').value = inquiryDates[inquiryDates.length - 1];
                // 日付絞り込み基準の初期設定を「問い合わせ日」にする
                document.getElementById('filter-date-type').value = 'inquiryDate';
            } else {
                // 日付が無い場合は「指定なし」にする
                document.getElementById('filter-date-type').value = 'none';
            }

            // フィルター項目の初期化と集計
            populateFilters(processedData);
            applyFiltersAndRender();

        } catch (error) {
            console.error(error);
            alert('Excelファイルの読み込み中にエラーが発生しました。\n\n【エラー詳細】\n' + error.message + '\n\n【スタックトレース】\n' + error.stack);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ヘッダー配列から同義語リストに基づき列インデックスを解決する
function resolveColumnMapping(headers) {
    const mapping = {};
    const normalizedHeaders = headers.map(h => h ? h.toString().trim().toLowerCase() : '');

    Object.keys(COLUMN_SYNONYMS).forEach(key => {
        const synonyms = COLUMN_SYNONYMS[key];
        let foundIdx = -1;

        // 1. 完全一致
        for (let syn of synonyms) {
            foundIdx = normalizedHeaders.indexOf(syn.toLowerCase());
            if (foundIdx !== -1) break;
        }

        // 2. 部分一致（完全一致で見つからなかった場合）
        if (foundIdx === -1) {
            for (let syn of synonyms) {
                // 1文字の同義語（「姓」「名」など）は誤爆しやすいため部分一致を除外
                if (syn.length <= 1) continue;
                foundIdx = normalizedHeaders.findIndex(h => h.includes(syn.toLowerCase()));
                if (foundIdx !== -1) break;
            }
        }

        mapping[key] = foundIdx; // 見つからなかった場合は -1
    });

    return mapping;
}

// 教室名の整形
function cleanStoreName(store) {
    if (!store) return '(空白)';
    return store.toString().trim();
}

// 日付フォーマットの統一
function formatDate(dateVal) {
    if (!dateVal) return '';
    if (dateVal instanceof Date) {
        return dateVal.toISOString().split('T')[0];
    }
    // 文字列日付のパース
    const dateStr = dateVal.toString().trim();
    const parsedDate = new Date(dateStr);
    if (!isNaN(parsedDate.getTime())) {
        return parsedDate.toISOString().split('T')[0];
    }
    return dateStr;
}

// ==========================================================================
// 2. 体験時学年の補正ロジック
// ==========================================================================
function calculateTrialGrade(currentGrade, dateStr, courseType = '', generation = '') {
    // 受講種別が「高校生以上」を含むか、または年代が20代以上（20代〜, 30代〜, ...）の場合は一律「一般・社会人」とする
    const isAdultCourse = courseType && courseType.includes('高校生以上');
    const isAdultGeneration = generation && /^[2-9]\d代/.test(generation);
    
    if (isAdultCourse || isAdultGeneration) {
        return '一般・社会人';
    }

    if (!currentGrade) return '(空白)';
    const gradeStr = currentGrade.toString().trim();
    
    // 遡り計算の対象外とする値
    if (gradeStr === '一般・社会人' || gradeStr === '未就学' || gradeStr === '年長') {
        return gradeStr;
    }

    // 体験日（または問合せ日）の年度（4月〜翌3月基準）を特定
    let targetYear = null;
    if (dateStr) {
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
            const year = date.getFullYear();
            const month = date.getMonth() + 1; // 0-indexedから1-indexedへ
            // 4月始まり
            targetYear = (month >= 4) ? year : (year - 1);
        }
    }

    if (targetYear === null) {
        return gradeStr; // 日付がパースできない場合は補正せず現在の学年を返す
    }

    // UIから基準日を取得（未設定時は本日）
    const baseDateInput = document.getElementById('base-date');
    let baseDate = new Date();
    if (baseDateInput && baseDateInput.value) {
        baseDate = new Date(baseDateInput.value);
    }
    const baseYearCalendar = baseDate.getFullYear();
    const baseMonth = baseDate.getMonth() + 1;
    const baseYear = (baseMonth >= 4) ? baseYearCalendar : (baseYearCalendar - 1); // 4月始まり基準
    
    const yearDiff = baseYear - targetYear;

    if (yearDiff <= 0) {
        return gradeStr; // 現在以降、または同年度の場合は補正なし
    }

    // 学年文字列のパース（例：「小学校3年」「中2」「高校1年」など）
    const match = gradeStr.match(/^(小学校|中学校|高校|中学|小学|高|中|小)(\d+)年$/);
    if (match) {
        const schoolType = match[1];
        const gradeNum = parseInt(match[2], 10);

        // 表記の正規化
        let schoolName = schoolType;
        if (schoolType === '小学' || schoolType === '小') schoolName = '小学校';
        if (schoolType === '中学' || schoolType === '中') schoolName = '中学校';
        if (schoolType === '高') schoolName = '高校';

        // 年度差を引き算する
        const adjustedGrade = gradeNum - yearDiff;

        if (adjustedGrade >= 1) {
            return `${schoolName}${adjustedGrade}年`;
        } else {
            // 小学校1年未満になった場合は、学校区分を遡らせる
            if (schoolName === '小学校') {
                return '年長・未就学';
            } else if (schoolName === '中学校') {
                // 中学未満は小学校へ (中学校1年 = 小学校7年相当)
                const elemGrade = 6 + adjustedGrade; // 1(中1) - 1(diff)=0 -> 小6
                return elemGrade >= 1 ? `小学校${elemGrade}年` : '年長・未就学';
            } else if (schoolName === '高校') {
                // 高校未満は中学校へ
                const jhsGrade = 3 + adjustedGrade; // 1(高1) - 1(diff)=0 -> 中3
                if (jhsGrade >= 1) {
                    return `中学校${jhsGrade}年`;
                } else {
                    const elemGrade = 6 + jhsGrade;
                    return elemGrade >= 1 ? `小学校${elemGrade}年` : '年長・未就学';
                }
            }
        }
    }

    return gradeStr; // パターンマッチしない場合はそのまま
}

// ==========================================================================
// 3. フィルターと集計ロジック
// ==========================================================================
function initFilters() {
    const filters = ['filter-store', 'filter-purpose', 'filter-grade', 'filter-status', 'filter-date-type', 'filter-date-start', 'filter-date-end'];
    filters.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', applyFiltersAndRender);
    });
}

// ユニーク値を抽出してフィルターセレクトボックスを生成
function populateFilters(data) {
    const stores = getUniqueSortedValues(data, 'store');
    const purposes = getUniqueSortedValues(data, 'purpose');
    const grades = getUniqueSortedValues(data, 'trialGrade', sortGrades);
    const statuses = getUniqueSortedValues(data, 'status');
    const courseTypes = getUniqueSortedValues(data, 'courseType');

    updateSelectOptions('filter-store', stores);
    updateSelectOptions('filter-purpose', purposes);
    updateSelectOptions('filter-grade', grades);
    updateSelectOptions('filter-status', statuses);

    const courseContainer = document.getElementById('filter-courseType-container');
    if (courseContainer) {
        courseContainer.innerHTML = '';
        courseTypes.forEach(val => {
            const label = document.createElement('label');
            label.style.display = 'block';
            label.style.cursor = 'pointer';
            label.style.marginBottom = '4px';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = val;
            cb.checked = true; // デフォルトで全てチェック
            cb.className = 'filter-courseType-cb';
            cb.style.marginRight = '6px';
            cb.addEventListener('change', applyFiltersAndRender);
            label.appendChild(cb);
            label.appendChild(document.createTextNode(val));
            courseContainer.appendChild(label);
        });
    }
}

function getUniqueSortedValues(data, key, sortFn = null) {
    const vals = data.map(item => item[key]).filter(v => v !== undefined && v !== '');
    const unique = [...new Set(vals)];
    return sortFn ? unique.sort(sortFn) : unique.sort((a, b) => a.localeCompare(b, 'ja'));
}

// 学年順のソートヘルパー
function sortGrades(a, b) {
    const order = ['年長・未就学', '年長', '未就学', 
                   '小学校1年', '小学校2年', '小学校3年', '小学校4年', '小学校5年', '小学校6年',
                   '中学校1年', '中学校2年', '中学校3年',
                   '高校1年', '高校2年', '高校3年', '一般・社会人', '(空白)'];
    
    const idxA = order.indexOf(a);
    const idxB = order.indexOf(b);
    
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b, 'ja');
}

function updateSelectOptions(id, values) {
    const select = document.getElementById(id);
    // すべてのオプションを削除（最初の「すべて」を残す）
    select.innerHTML = '<option value="all">すべて</option>';
    values.forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        select.appendChild(opt);
    });
}

// フィルターを適用し、画面全体を再レンダリングする
function applyFiltersAndRender() {
    const storeVal = document.getElementById('filter-store').value;
    const purposeVal = document.getElementById('filter-purpose').value;
    const gradeVal = document.getElementById('filter-grade').value;
    const statusVal = document.getElementById('filter-status').value;
    const checkedCourseCbs = Array.from(document.querySelectorAll('.filter-courseType-cb:checked')).map(cb => cb.value);

    // 日付絞り込みの取得
    const dateType = document.getElementById('filter-date-type').value;
    const dateStart = document.getElementById('filter-date-start').value;
    const dateEnd = document.getElementById('filter-date-end').value;

    allStoreFilteredData = processedData.filter(item => {
        const matchPurpose = (purposeVal === 'all' || item.purpose === purposeVal);
        const matchGrade = (gradeVal === 'all' || item.trialGrade === gradeVal);
        const matchStatus = (statusVal === 'all' || item.status === statusVal);
        const matchCourse = checkedCourseCbs.length === 0 || checkedCourseCbs.includes(item.courseType);

        // 日付範囲の判定
        let matchDate = true;
        if (dateType !== 'none') {
            const dateStr = item[dateType + 'Str']; // e.g. item.inquiryDateStr
            if (!dateStr) {
                matchDate = false; // 日付データが無い場合は除外
            } else {
                if (dateStart && dateStr < dateStart) matchDate = false;
                if (dateEnd && dateStr > dateEnd) matchDate = false;
            }
        }

        return matchPurpose && matchGrade && matchStatus && matchCourse && matchDate;
    });

    filteredData = allStoreFilteredData.filter(item => {
        return (storeVal === 'all' || item.store === storeVal);
    });

    renderKPIs(filteredData);
    renderCharts(filteredData);
    renderStoreTable(allStoreFilteredData); // 全教室KPI比較マトリクス用（教室フィルタを無視）
    renderDetailTable(filteredData);
    
    // ファネル分析コメント（インサイト）の生成と表示
    renderFunnelInsights(filteredData, allStoreFilteredData, storeVal, checkedCourseCbs);
}

// KPIカードの集計・表示
function renderKPIs(data) {
    const total = data.length;
    
    // 体験実施数 (体験キャンセル、体験予約に至らず、問合せ を除く)
    const trialData = data.filter(item => !['問合せ', '体験キャンセル', '体験予約に至らず'].includes(item.status));
    const trialCount = trialData.length;
    const trialRate = total > 0 ? (trialCount / total * 100).toFixed(1) : 0;
    
    // 入会者（受講生 ＋ 退会生）の定義
    const enrollData = data.filter(item => item.status === '受講生' || item.status === '退会生');
    const enrollCount = enrollData.length;
    
    // 体験した上で未入会（NG・持ち帰り等）になった数
    const ngCount = data.filter(item => ['NG', '持ち帰り', '入会キャンセル'].includes(item.status)).length;

    // 体験した方（入会 ＋ 未入会）を分母とする
    const closedCount = enrollCount + ngCount;
    
    // 体験した方の入会率
    const trialEnrollRate = closedCount > 0 ? (enrollCount / closedCount * 100).toFixed(1) : 0;
    // 問合せからの入会率
    const totalEnrollRate = total > 0 ? (enrollCount / total * 100).toFixed(1) : 0;

    // 体験した方の未入会率
    const trialNgRate = closedCount > 0 ? (ngCount / closedCount * 100).toFixed(1) : 0;

    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-trial').innerHTML = `${trialCount.toLocaleString()} <span style="font-size: 1.2rem; font-weight: normal; color: var(--text-secondary);">(${trialRate}%)</span>`;
    document.getElementById('kpi-enroll').textContent = enrollCount.toLocaleString();
    document.getElementById('kpi-trial-enroll-rate').textContent = `${trialEnrollRate}%`;
    document.getElementById('kpi-total-enroll-rate').textContent = `${totalEnrollRate}%`;
    document.getElementById('kpi-trial-ng-rate').textContent = `${trialNgRate}%`;
}

// ファネル分析インサイトパネルの生成と表示
function renderFunnelInsights(storeData, allStoreData, storeVal, checkedCourseCbs) {
    const insightPanel = document.getElementById('insight-panel');
    if (!insightPanel) return;

    if (storeVal === 'all' || storeData.length === 0 || allStoreData.length === 0) {
        insightPanel.style.display = 'none';
        return;
    }

    // 計算用ヘルパー
    const calcRates = (data) => {
        const total = data.length;
        const trialData = data.filter(item => !['問合せ', '体験キャンセル', '体験予約に至らず'].includes(item.status));
        const trialCount = trialData.length;
        const trialRate = total > 0 ? (trialCount / total) * 100 : 0;
        
        const enrollData = data.filter(item => item.status === '受講生' || item.status === '退会生');
        const enrollCount = enrollData.length;
        const ngCount = data.filter(item => ['NG', '持ち帰り', '入会キャンセル'].includes(item.status)).length;
        const closedCount = enrollCount + ngCount;
        
        const trialEnrollRate = closedCount > 0 ? (enrollCount / closedCount) * 100 : 0;
        
        return { trialRate, trialEnrollRate, total };
    };

    // 選択された受講種別を特定（未選択の場合はデータに存在する全受講種別を取得）
    let targetCourses = checkedCourseCbs;
    if (!targetCourses || targetCourses.length === 0) {
        targetCourses = [...new Set(storeData.map(item => item.courseType))].filter(Boolean);
    }

    let insightHtml = `<div class="insight-title">💡 ${storeVal} のファネル分析・課題検知</div>`;

    targetCourses.forEach(course => {
        const courseStoreData = storeData.filter(item => item.courseType === course);
        const courseAllData = allStoreData.filter(item => item.courseType === course);
        
        if (courseStoreData.length === 0) return; // 対象データがない場合はスキップ

        const storeRates = calcRates(courseStoreData);
        const allRates = calcRates(courseAllData);

        const storeCourseKey = `${storeVal}::${course}`;
        const foundSummary = currentStoreSummaryList.find(s => `${s.name}::${s.courseType}` === storeCourseKey);
        
        let rankHtml = '';
        if (foundSummary) {
            rankHtml = ` <span class="rank-badge ${foundSummary.rankClass}" style="transform: scale(0.85); display: inline-block; vertical-align: middle; margin-top:-3px;">${foundSummary.rank}</span>`;
        }

        insightHtml += `<div class="insight-body" style="margin-bottom: 15px; border-left: 3px solid var(--border); padding-left: 10px;">`;
        insightHtml += `<h4 style="margin: 0 0 10px 0; color: var(--text-primary);">【受講種別: ${course}${rankHtml}】 (サンプル数: ${storeRates.total}件)</h4>`;
        
        let hasWarning = false;

        // 全体の店舗数を算出（平均計算用）
        const uniqueStoresCount = new Set(allStoreData.map(item => item.store)).size || 1;
        const averageInquiries = allRates.total / uniqueStoresCount;
        const averageInquiriesFormatted = averageInquiries.toFixed(1);

        // 問合せ数の比較
        if (storeRates.total < averageInquiries * 0.8) {
            insightHtml += `
                <div style="margin-bottom: 12px;">
                    <p class="insight-alert" style="margin-bottom:0;">⚠️ <strong>問合せ数が少ないです (${storeRates.total}件)</strong><br>1店舗あたりの全体平均(${averageInquiriesFormatted}件)を下回っています。このコースの集客・認知活動に課題がある可能性があります。</p>
                    <button class="btn-hearing-guide" onclick="toggleHearingGuide(this)">📋 現場ヒアリング項目を確認</button>
                    <div class="hearing-guide-box" style="display: none;">
                        <div class="hearing-guide-title">💡 想定されるボトルネック（仮説）</div>
                        <p style="margin-bottom: 8px; color: var(--text-secondary);">地域認知度低下、近近の競合校の新規出店、開講曜日・時間帯の競合など。</p>
                        <div class="hearing-guide-title">📋 ヒアリングチェックリスト</div>
                        <ul class="hearing-guide-list">
                            <li>近隣の競合状況の変化（チラシや看板など）は見られますか？</li>
                            <li>直近でポスティングや門前配布、紹介キャンペーンなどの募集活動を実施できていますか？</li>
                            <li>ターゲット層の通いづらい曜日や時間帯に開講していませんか？</li>
                        </ul>
                    </div>
                </div>
            `;
            hasWarning = true;
        } else if (storeRates.total > averageInquiries * 1.2) {
            insightHtml += `<p class="insight-success">✅ <strong>問合せ数が豊富です (${storeRates.total}件)</strong><br>1店舗あたりの全体平均(${averageInquiriesFormatted}件)を上回っており、集客が非常に好調です。</p>`;
        } else {
            insightHtml += `<p class="insight-neutral">ℹ️ 問合せ数 (${storeRates.total}件) は1店舗あたりの全体平均(${averageInquiriesFormatted}件)と同等水準です。</p>`;
        }

        // 体験率の比較
        const trialDiff = storeRates.trialRate - allRates.trialRate;
        if (trialDiff < -5) {
            const isAdult = course.includes('高校生以上');
            const bottleneckText = isAdult
                ? '初期対応スピードの遅れ、大人の生活サイクルに合わせた体験枠の不足、初心者向けの安心感（敷居の低さ）の訴求不足など。'
                : '初期対応スピードの遅れ、案内時の体験メリットの訴求不足、体験枠不足など。';
                
            const listHtml = isAdult
                ? `<li>問い合わせから初回連絡まで何時間（または何日）で対応できていますか？（大人の活動時間帯に即対応できているか）</li>
                   <li>案内メール・電話で本人の受講目的（仕事、趣味等）への共感や、初心者でも安心できる案内ができていますか？</li>
                   <li>体験可能な枠（仕事帰りの時間帯や土日など）が不足していませんか？</li>`
                : `<li>問合せが入ってから最初の連絡まで何時間（または何日）で対応できていますか？（即日対応できているか）</li>
                   <li>案内メール・電話での最初の連絡の文面やトーク内容は適切ですか？</li>
                   <li>体験可能な枠（特に土日など保護者の希望枠）が不足していませんか？</li>`;

            insightHtml += `
                <div style="margin-bottom: 12px;">
                    <p class="insight-alert" style="margin-bottom:0;">⚠️ <strong>体験実施率が低いです (${storeRates.trialRate.toFixed(1)}%)</strong><br>全体平均(${allRates.trialRate.toFixed(1)}%)を大きく下回っています。電話・メール対応スピード、体験受け入れ体制、案内トークに課題がある可能性があります。</p>
                    <button class="btn-hearing-guide" onclick="toggleHearingGuide(this)">📋 現場ヒアリング項目を確認</button>
                    <div class="hearing-guide-box" style="display: none;">
                        <div class="hearing-guide-title">💡 想定されるボトルネック（仮説）</div>
                        <p style="margin-bottom: 8px; color: var(--text-secondary);">${bottleneckText}</p>
                        <div class="hearing-guide-title">📋 ヒアリングチェックリスト</div>
                        <ul class="hearing-guide-list">
                            ${listHtml}
                        </ul>
                    </div>
                </div>
            `;
            hasWarning = true;
        } else if (trialDiff > 5) {
            insightHtml += `<p class="insight-success">✅ <strong>体験実施率が優秀です (${storeRates.trialRate.toFixed(1)}%)</strong><br>全体平均(${allRates.trialRate.toFixed(1)}%)を上回っており、集客から体験への誘導が非常にうまくいっています。</p>`;
        } else {
            insightHtml += `<p class="insight-neutral">ℹ️ 体験実施率 (${storeRates.trialRate.toFixed(1)}%) は全体平均(${allRates.trialRate.toFixed(1)}%)と同等水準です。</p>`;
        }

        // 入会率の比較
        const enrollDiff = storeRates.trialEnrollRate - allRates.trialEnrollRate;
        if (enrollDiff < -5) {
            const isAdult = course.includes('高校生以上');
            const bottleneckText = isAdult
                ? '体験授業の難易度ミスマッチ（難しすぎる/簡単すぎる）、受講目的（実務、趣味、資格など）に直結するメリットの提示不足、クロージングや持ち帰り後のフォロー不足など。'
                : '体験授業の満足度不足、お子様の興味に合わないコース提案、体験面談でのクロージングや追客フォロー不足など。';
                
            const listHtml = isAdult
                ? `<li>体験授業時、受講者本人の「やりたいこと」に寄り添った内容を提供でき、達成感を感じてもらえましたか？</li>
                   <li>面談時に本人の目的（仕事、資格、趣味など）に合わせたコース提案や、無理のないスケジュール（振替制度など）の案内ができましたか？</li>
                   <li>「検討します」と持ち帰った本人に対し、その後のフォロー連絡（追客）を行っていますか？</li>`
                : `<li>体験授業時のお子様の反応や作ったものに満足感はありましたか？</li>
                   <li>体験面談時に保護者へどのようなコースを提案し、どんな懸念点を言われましたか？</li>
                   <li>「検討します」と持ち帰った保護者へ、その後のフォロー連絡（追客）を行っていますか？</li>`;

            insightHtml += `
                <div style="margin-bottom: 12px;">
                    <p class="insight-alert" style="margin-bottom:0;">⚠️ <strong>体験入会率に課題があります (${storeRates.trialEnrollRate.toFixed(1)}%)</strong><br>全体平均(${allRates.trialEnrollRate.toFixed(1)}%)を下回っています。体験授業の質、コース提案、クロージングトークの強化が必要です。</p>
                    <button class="btn-hearing-guide" onclick="toggleHearingGuide(this)">📋 現場ヒアリング項目を確認</button>
                    <div class="hearing-guide-box" style="display: none;">
                        <div class="hearing-guide-title">💡 想定されるボトルネック（仮説）</div>
                        <p style="margin-bottom: 8px; color: var(--text-secondary);">${bottleneckText}</p>
                        <div class="hearing-guide-title">📋 ヒアリングチェックリスト</div>
                        <ul class="hearing-guide-list">
                            ${listHtml}
                        </ul>
                    </div>
                </div>
            `;
            hasWarning = true;
        } else if (enrollDiff > 5) {
            insightHtml += `<p class="insight-success">✅ <strong>体験入会率が優秀です (${storeRates.trialEnrollRate.toFixed(1)}%)</strong><br>全体平均(${allRates.trialEnrollRate.toFixed(1)}%)を上回っています。現場のクロージング力が高い状態です。</p>`;
        } else {
            insightHtml += `<p class="insight-neutral">ℹ️ 体験入会率 (${storeRates.trialEnrollRate.toFixed(1)}%) は全体平均(${allRates.trialEnrollRate.toFixed(1)}%)と同等水準です。</p>`;
        }

        // 総括コメント
        if (!hasWarning && storeRates.trialRate > allRates.trialRate && storeRates.trialEnrollRate > allRates.trialEnrollRate) {
            insightHtml += `<p class="insight-summary">🎯 <strong>総括:</strong> 非常に優秀な歩留まりです。成功要因の横展開をお勧めします。</p>`;
        } else if (hasWarning) {
            insightHtml += `<p class="insight-summary-alert">🚨 <strong>総括:</strong> アラートが出ている指標について、現場へのヒアリングやオペレーション確認を行ってください。</p>`;
        }
        
        insightHtml += `</div>`;
    });

    insightPanel.innerHTML = insightHtml;
    insightPanel.style.display = 'block';
}

// ==========================================================================
// 4. グラフ描画 (Chart.js)
// ==========================================================================
function destroyCharts() {
    Object.keys(charts).forEach(key => {
        if (charts[key]) {
            charts[key].destroy();
            charts[key] = null;
        }
    });
}

function renderCharts(data) {
    destroyCharts();

    // --- A. 学年別チャートのデータ集計 ---
    const grades = getUniqueSortedValues(processedData, 'trialGrade', sortGrades); // 常に全学年の並び順を固定
    const gradeCounts = {};
    const gradeEnrolls = {};
    const gradeNgs = {};

    grades.forEach(g => {
        gradeCounts[g] = 0;
        gradeEnrolls[g] = 0;
        gradeNgs[g] = 0;
    });

    data.forEach(item => {
        if (gradeCounts[item.trialGrade] !== undefined) {
            gradeCounts[item.trialGrade]++;
            if (item.status === '受講生' || item.status === '退会生') {
                gradeEnrolls[item.trialGrade]++;
            } else if (['NG', '持ち帰り', '入会キャンセル'].includes(item.status)) {
                gradeNgs[item.trialGrade]++;
            }
        }
    });

    const gradeLabels = grades.filter(g => gradeCounts[g] > 0); // データがある学年のみ表示
    const enrollData = gradeLabels.map(g => gradeEnrolls[g]);
    const ngData = gradeLabels.map(g => gradeNgs[g]);
    const otherData = gradeLabels.map(g => gradeCounts[g] - gradeEnrolls[g] - gradeNgs[g]);
    const rateData = gradeLabels.map(g => {
        const closed = gradeEnrolls[g] + gradeNgs[g];
        return closed > 0 ? (gradeEnrolls[g] / closed * 100).toFixed(1) : 0;
    });

    // 学年別テーブルの描画
    const gradeTbody = document.querySelector('#chart-grade-table tbody');
    if (gradeTbody) {
        gradeTbody.innerHTML = '';
        gradeLabels.forEach((g, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${g}</strong></td>
                <td>${enrollData[idx]}</td>
                <td>${ngData[idx]}</td>
                <td>${otherData[idx]}</td>
                <td>${gradeCounts[g]}</td>
                <td><strong>${rateData[idx]}%</strong></td>
            `;
            gradeTbody.appendChild(tr);
        });
    }

    // 平均体験入会率の算出と偏りのある箇所のハイライト判定
    const totalGradeEnroll = enrollData.reduce((a, b) => a + b, 0);
    const totalGradeNg = ngData.reduce((a, b) => a + b, 0);
    const totalGradeClosed = totalGradeEnroll + totalGradeNg;
    const avgGradeRate = totalGradeClosed > 0 ? (totalGradeEnroll / totalGradeClosed * 100) : 0;

    const gradeBarBackgrounds = enrollData.map((val, idx) => {
        const rate = parseFloat(rateData[idx]);
        if (avgGradeRate > 0 && rate < avgGradeRate - 10) {
            return 'rgba(220, 38, 38, 0.9)'; // 偏差が10%以上低い場合は赤色警告
        }
        return 'rgba(244, 63, 94, 0.75)'; // 通常のローズ
    });

    const gradeLinePointColors = rateData.map((rateStr, idx) => {
        const rate = parseFloat(rateStr);
        if (avgGradeRate > 0 && rate < avgGradeRate - 10) {
            return 'rgb(220, 38, 38)'; // 警告赤
        }
        return 'rgb(244, 63, 94)'; // 通常ローズ
    });

    const gradeLinePointRadius = rateData.map((rateStr, idx) => {
        const rate = parseFloat(rateStr);
        if (avgGradeRate > 0 && rate < avgGradeRate - 10) {
            return 7; // ドットを大きく
        }
        return 3;
    });

    // 学年別グラフ描画（混合チャート）
    const ctxGrade = document.getElementById('chart-grade').getContext('2d');
    charts.grade = new Chart(ctxGrade, {
        type: 'bar',
        data: {
            labels: gradeLabels,
            datasets: [
                {
                    label: '入会',
                    data: enrollData,
                    backgroundColor: gradeBarBackgrounds,
                    borderColor: 'rgb(244, 63, 94)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'NG・持ち帰り',
                    data: ngData,
                    backgroundColor: 'rgba(245, 158, 11, 0.75)',
                    borderColor: 'rgb(245, 158, 11)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'その他 (問合せ等)',
                    data: otherData,
                    backgroundColor: 'rgba(99, 102, 241, 0.4)',
                    borderColor: 'rgba(99, 102, 241, 0.7)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: '体験した方の入会率 (%)',
                    data: rateData,
                    type: 'line',
                    borderColor: 'rgb(244, 63, 94)',
                    backgroundColor: 'rgb(244, 63, 94)',
                    pointBackgroundColor: gradeLinePointColors,
                    pointBorderColor: gradeLinePointColors,
                    pointRadius: gradeLinePointRadius,
                    pointHoverRadius: gradeLinePointRadius.map(r => r + 2),
                    borderWidth: 3,
                    fill: false,
                    tension: 0.2,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    stacked: true,
                    title: { display: true, text: '問合せ数 (件)' },
                    grid: { drawOnChartArea: true }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    title: { display: true, text: '入会率 (%)' },
                    min: 0,
                    max: 100,
                    grid: { drawOnChartArea: false }
                }
            }
        }
    });

    // --- B. 来校目的別チャートのデータ集計 ---
    const purposes = getUniqueSortedValues(processedData, 'purpose');
    const purposeCounts = {};
    const purposeEnrolls = {};
    const purposeNgs = {};

    purposes.forEach(p => {
        purposeCounts[p] = 0;
        purposeEnrolls[p] = 0;
        purposeNgs[p] = 0;
    });

    data.forEach(item => {
        if (purposeCounts[item.purpose] !== undefined) {
            purposeCounts[item.purpose]++;
            if (item.status === '受講生' || item.status === '退会生') {
                purposeEnrolls[item.purpose]++;
            } else if (['NG', '持ち帰り', '入会キャンセル'].includes(item.status)) {
                purposeNgs[item.purpose]++;
            }
        }
    });

    const purposeLabels = purposes.filter(p => purposeCounts[p] > 0);
    const purposeEnrollData = purposeLabels.map(p => purposeEnrolls[p]);
    const purposeNgData = purposeLabels.map(p => purposeNgs[p]);
    const purposeOtherData = purposeLabels.map(p => purposeCounts[p] - purposeEnrolls[p] - purposeNgs[p]);
    const purposeRateData = purposeLabels.map(p => {
        const closed = purposeEnrolls[p] + purposeNgs[p];
        return closed > 0 ? (purposeEnrolls[p] / closed * 100).toFixed(1) : 0;
    });

    // 目的別テーブルの描画
    const purposeTbody = document.querySelector('#chart-purpose-table tbody');
    if (purposeTbody) {
        purposeTbody.innerHTML = '';
        purposeLabels.forEach((p, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${p}</strong></td>
                <td>${purposeEnrollData[idx]}</td>
                <td>${purposeNgData[idx]}</td>
                <td>${purposeOtherData[idx]}</td>
                <td>${purposeCounts[p]}</td>
                <td><strong>${purposeRateData[idx]}%</strong></td>
            `;
            purposeTbody.appendChild(tr);
        });
    }

    // 平均体験入会率の算出と偏りのある箇所のハイライト判定
    const totalPurposeEnroll = purposeEnrollData.reduce((a, b) => a + b, 0);
    const totalPurposeNg = purposeNgData.reduce((a, b) => a + b, 0);
    const totalPurposeClosed = totalPurposeEnroll + totalPurposeNg;
    const avgPurposeRate = totalPurposeClosed > 0 ? (totalPurposeEnroll / totalPurposeClosed * 100) : 0;

    const purposeBarBackgrounds = purposeEnrollData.map((val, idx) => {
        const rate = parseFloat(purposeRateData[idx]);
        if (avgPurposeRate > 0 && rate < avgPurposeRate - 10) {
            return 'rgba(220, 38, 38, 0.9)'; // 偏差が10%以上低い場合は赤色警告
        }
        return 'rgba(244, 63, 94, 0.75)'; // 通常ローズ
    });

    const purposeLinePointColors = purposeRateData.map((rateStr, idx) => {
        const rate = parseFloat(rateStr);
        if (avgPurposeRate > 0 && rate < avgPurposeRate - 10) {
            return 'rgb(220, 38, 38)'; // 警告赤
        }
        return 'rgb(244, 63, 94)'; // 通常ローズ
    });

    const purposeLinePointRadius = purposeRateData.map((rateStr, idx) => {
        const rate = parseFloat(rateStr);
        if (avgPurposeRate > 0 && rate < avgPurposeRate - 10) {
            return 7; // ドットを大きく
        }
        return 3;
    });

    // 来校目的別グラフ描画（混合チャート）
    const ctxPurpose = document.getElementById('chart-purpose').getContext('2d');
    charts.purpose = new Chart(ctxPurpose, {
        type: 'bar',
        data: {
            labels: purposeLabels,
            datasets: [
                {
                    label: '入会',
                    data: purposeEnrollData,
                    backgroundColor: purposeBarBackgrounds,
                    borderColor: 'rgb(244, 63, 94)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'NG・持ち帰り',
                    data: purposeNgData,
                    backgroundColor: 'rgba(245, 158, 11, 0.75)',
                    borderColor: 'rgb(245, 158, 11)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'その他 (問合せ等)',
                    data: purposeOtherData,
                    backgroundColor: 'rgba(6, 182, 212, 0.4)',
                    borderColor: 'rgba(6, 182, 212, 0.7)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: '体験した方の入会率 (%)',
                    data: purposeRateData,
                    type: 'line',
                    borderColor: 'rgb(244, 63, 94)',
                    backgroundColor: 'rgb(244, 63, 94)',
                    pointBackgroundColor: purposeLinePointColors,
                    pointBorderColor: purposeLinePointColors,
                    pointRadius: purposeLinePointRadius,
                    pointHoverRadius: purposeLinePointRadius.map(r => r + 2),
                    borderWidth: 3,
                    fill: false,
                    tension: 0.2,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    stacked: true,
                    title: { display: true, text: '問合せ数 (件)' }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    title: { display: true, text: '入会率 (%)' },
                    min: 0,
                    max: 100,
                    grid: { drawOnChartArea: false }
                }
            }
        }
    });

    // --- C. 前年比較チャートのデータ集計 ---
    // 基準日から今年と前年を判定
    const baseDateInput = document.getElementById('base-date');
    let baseDateVal = new Date();
    if (baseDateInput && baseDateInput.value) {
        baseDateVal = new Date(baseDateInput.value);
    }

    // processedData から抽出して全期間の年リストを作成
    const getYearFromDate = (dateStr) => {
        if (!dateStr || dateStr === '(空白)') return null;
        const parts = dateStr.split('-');
        if (parts.length >= 1) {
            const y = parseInt(parts[0], 10);
            if (!isNaN(y)) return y;
        }
        return null;
    };

    const allYearsSet = new Set();
    processedData.forEach(item => {
        const y = getYearFromDate(item.inquiryDateStr);
        if (y) allYearsSet.add(y);
    });
    const sortedYears = [...allYearsSet].sort((a, b) => b - a); // 降順ソート（直近の年が先頭）

    // UIのタイトルや見出しを動的に更新
    const yoyInqTitle = document.getElementById('yoy-inq-title');
    const yoyRateTitle = document.getElementById('yoy-rate-title');
    const displayModeText = yoyMode === 'inq' ? '問合せ数' : '入会者数';
    if (yoyInqTitle) yoyInqTitle.textContent = `${displayModeText} 経年推移 (${sortedYears.join(' vs ')})`;
    if (yoyRateTitle) yoyRateTitle.textContent = `体験入会率 経年推移 (${sortedYears.join(' vs ')})`;

    // 動的にヘッダーを生成
    const tableHeader = document.querySelector('#chart-yoy-table thead tr');
    if (tableHeader) {
        let html = '<th>月</th>';
        sortedYears.forEach(y => { html += `<th>${y}年 問合せ</th>`; });
        sortedYears.forEach(y => { html += `<th>${y}年 入会</th>`; });
        sortedYears.forEach(y => { html += `<th>${y}年 入会率</th>`; });
        tableHeader.innerHTML = html;
    }

    // 1月〜12月の配列を作成
    const months = Array.from({length: 12}, (_, i) => i + 1); // 1〜12
    const monthLabels = months.map(m => `${m}月`);

    // 各年ごとの各月の集計オブジェクトを初期化
    const yoyData = {};
    sortedYears.forEach(y => {
        yoyData[y] = {
            inq: Array(12).fill(0),
            enroll: Array(12).fill(0),
            closed: Array(12).fill(0),
            rate: Array(12).fill(0)
        };
    });

    data.forEach(item => {
        if (!item.inquiryDateStr) return;
        const parts = item.inquiryDateStr.split('-');
        if (parts.length < 2) return;
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10); // 1〜12

        if (month >= 1 && month <= 12 && yoyData[year]) {
            const mIdx = month - 1;
            const isEnroll = (item.status === '受講生' || item.status === '退会生');
            const isClosed = isEnroll || ['NG', '持ち帰り', '入会キャンセル'].includes(item.status);

            yoyData[year].inq[mIdx]++;
            if (isEnroll) yoyData[year].enroll[mIdx]++;
            if (isClosed) yoyData[year].closed[mIdx]++;
        }
    });

    // 割合の計算
    sortedYears.forEach(y => {
        for (let i = 0; i < 12; i++) {
            const closed = yoyData[y].closed[i];
            yoyData[y].rate[i] = closed > 0 ? (yoyData[y].enroll[i] / closed * 100).toFixed(1) : 0;
        }
    });

    // YoY比較テーブルの描画
    const yoyTbody = document.querySelector('#chart-yoy-table tbody');
    if (yoyTbody) {
        yoyTbody.innerHTML = '';
        months.forEach((m, i) => {
            const tr = document.createElement('tr');
            let rowHtml = `<td><strong>${m}月</strong></td>`;
            // Inquiries
            sortedYears.forEach(y => { rowHtml += `<td>${yoyData[y].inq[i]}</td>`; });
            // Enrolls
            sortedYears.forEach(y => { rowHtml += `<td>${yoyData[y].enroll[i]}</td>`; });
            // Rates
            sortedYears.forEach(y => { rowHtml += `<td><strong>${yoyData[y].rate[i]}%</strong></td>`; });
            tr.innerHTML = rowHtml;
            yoyTbody.appendChild(tr);
        });
    }

    // 配色用カラーパレット (直近の年から順に割り当てる)
    const yearColors = [
        'rgb(59, 130, 246)',  // Blue
        'rgb(244, 63, 94)',   // Rose
        'rgb(16, 185, 129)',  // Emerald
        'rgb(245, 158, 11)',  // Amber
        'rgb(139, 92, 246)',  // Violet
        'rgb(6, 182, 212)'    // Cyan
    ];

    // チャート1: 問合せ数 または 入会数の経年比較 (折れ線グラフ)
    const volumeDatasets = sortedYears.map((y, idx) => {
        const color = yearColors[idx % yearColors.length];
        const isCurrentYear = (idx === 0 && sortedYears.length > 1);
        
        let pointBg = color;
        let pointBorder = color;
        let pointRad = 3;
        
        if (isCurrentYear) {
            const y1 = sortedYears[1];
            const currentData = yoyMode === 'inq' ? yoyData[y].inq : yoyData[y].enroll;
            const prevData = yoyMode === 'inq' ? yoyData[y1].inq : yoyData[y1].enroll;
            
            pointBg = [];
            pointBorder = [];
            pointRad = [];
            
            for (let i = 0; i < 12; i++) {
                const currentVal = currentData[i];
                const prevVal = prevData[i];
                // 前年比で20%以上減少しているか判定
                const isDrop = prevVal > 0 && ((prevVal - currentVal) / prevVal) >= 0.20;
                
                if (isDrop) {
                    pointBg.push('rgb(220, 38, 38)'); // 警告用赤
                    pointBorder.push('rgb(185, 28, 28)');
                    pointRad.push(7);
                } else {
                    pointBg.push(color);
                    pointBorder.push(color);
                    pointRad.push(3);
                }
            }
        }

        return {
            label: `${y}年 ${yoyMode === 'inq' ? '問合せ' : '入会'}`,
            data: yoyMode === 'inq' ? yoyData[y].inq : yoyData[y].enroll,
            borderColor: color,
            backgroundColor: color.replace('rgb', 'rgba').replace(')', ', 0.1)'),
            pointBackgroundColor: pointBg,
            pointBorderColor: pointBorder,
            pointRadius: pointRad,
            pointHoverRadius: Array.isArray(pointRad) ? pointRad.map(r => r + 2) : 5,
            borderWidth: 2,
            fill: false,
            tension: 0.1
        };
    });

    const ctxYoyInq = document.getElementById('chart-yoy-inq').getContext('2d');
    charts.yoyInq = new Chart(ctxYoyInq, {
        type: 'line',
        data: {
            labels: monthLabels,
            datasets: volumeDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    title: { display: true, text: '件数' }
                }
            }
        }
    });

    // チャート2: 体験入会率の経年比較 (折れ線グラフ)
    const rateDatasets = sortedYears.map((y, idx) => {
        const color = yearColors[idx % yearColors.length];
        const isCurrentYear = (idx === 0 && sortedYears.length > 1);
        
        let pointBg = color;
        let pointBorder = color;
        let pointRad = 3;
        
        if (isCurrentYear) {
            const y1 = sortedYears[1];
            const currentData = yoyData[y].rate;
            const prevData = yoyData[y1].rate;
            
            pointBg = [];
            pointBorder = [];
            pointRad = [];
            
            for (let i = 0; i < 12; i++) {
                const currentVal = parseFloat(currentData[i]);
                const prevVal = parseFloat(prevData[i]);
                // 前年同期比で入会率が20%以上低下しているか判定
                const isDrop = prevVal > 0 && ((prevVal - currentVal) / prevVal) >= 0.20;
                
                if (isDrop) {
                    pointBg.push('rgb(220, 38, 38)'); // 警告用赤
                    pointBorder.push('rgb(185, 28, 28)');
                    pointRad.push(7);
                } else {
                    pointBg.push(color);
                    pointBorder.push(color);
                    pointRad.push(3);
                }
            }
        }

        return {
            label: `${y}年 入会率`,
            data: yoyData[y].rate,
            borderColor: color,
            backgroundColor: color.replace('rgb', 'rgba').replace(')', ', 0.1)'),
            pointBackgroundColor: pointBg,
            pointBorderColor: pointBorder,
            pointRadius: pointRad,
            pointHoverRadius: Array.isArray(pointRad) ? pointRad.map(r => r + 2) : 5,
            borderWidth: 2.5,
            fill: false,
            tension: 0.1
        };
    });

    const ctxYoyRate = document.getElementById('chart-yoy-rate').getContext('2d');
    charts.yoyRate = new Chart(ctxYoyRate, {
        type: 'line',
        data: {
            labels: monthLabels,
            datasets: rateDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    title: { display: true, text: '入会率 (%)' },
                    min: 0,
                    max: 100
                }
            }
        }
    });
}

// ==========================================================================
// 5. テーブル描画
// ==========================================================================

// 教室（個店）別集計テーブル
function renderStoreTable(data) {
    const tbody = document.querySelector('#store-table tbody');
    tbody.innerHTML = '';

    // 教室×受講種別ごとに集計
    const storeSummary = {};
    data.forEach(item => {
        const course = item.courseType || '未設定';
        const key = `${item.store}::${course}`;
        
        if (!storeSummary[key]) {
            storeSummary[key] = {
                name: item.store,
                courseType: course,
                total: 0,
                trialDone: 0,
                enroll: 0,
                ng: 0,
                withdraw: 0,
                grades: {} // 学年ごとの集計用
            };
        }

        const summary = storeSummary[key];
        summary.total++;
        
        // 学年集計オブジェクトの初期化
        const gradeKey = item.trialGrade || '不明';
        if (!summary.grades[gradeKey]) {
            summary.grades[gradeKey] = { trialDone: 0, ng: 0, enroll: 0 };
        }

        // 体験実施数 (体験キャンセル、体験予約に至らず、問合せ を除く)
        const isNotTrial = ['問合せ', '体験キャンセル', '体験予約に至らず'].includes(item.status);
        if (!isNotTrial) {
            summary.trialDone++;
            summary.grades[gradeKey].trialDone++;
        }

        if (item.status === '受講生' || item.status === '退会生') {
            summary.enroll++;
            summary.grades[gradeKey].enroll++;
        }
        if (['NG', '持ち帰り', '入会キャンセル'].includes(item.status)) {
            summary.ng++;
            summary.grades[gradeKey].ng++;
        }
        if (item.status === '退会生') {
            summary.withdraw++;
        }
    });

    // 全体の平均値を算出（ランク評価用）
    let totalInq = 0, totalTrial = 0, totalEnr = 0, totalClosedAll = 0;
    const storeKeys = Object.keys(storeSummary);
    const storeCount = storeKeys.length || 1;
    storeKeys.forEach(key => {
        totalInq += storeSummary[key].total;
        totalTrial += storeSummary[key].trialDone;
        totalEnr += storeSummary[key].enroll;
        totalClosedAll += (storeSummary[key].enroll + storeSummary[key].ng);
    });
    const avgInq = totalInq / storeCount;
    const avgTrial = totalInq > 0 ? (totalTrial / totalInq * 100) : 0;
    const avgEnrRate = totalClosedAll > 0 ? (totalEnr / totalClosedAll * 100) : 0;
    const avgEnrCount = totalEnr / storeCount;

    // 各教室の割合データを算出してソート用のリストを作成
    const summaryList = Object.values(storeSummary).map(s => {
        const closedCount = s.enroll + s.ng; // A方式の分母
        s.trialEnrollRate = closedCount > 0 ? parseFloat((s.enroll / closedCount * 100).toFixed(1)) : 0;
        s.trialNgRate = closedCount > 0 ? parseFloat((s.ng / closedCount * 100).toFixed(1)) : 0;
        s.totalEnrollRate = s.total > 0 ? parseFloat((s.enroll / s.total * 100).toFixed(1)) : 0;
        s.totalNgRate = s.total > 0 ? parseFloat((s.ng / s.total * 100).toFixed(1)) : 0;
        
        // 総合ランクの算出
        const sTrialRate = s.total > 0 ? (s.trialDone / s.total * 100) : 0;
        const inqScore = avgInq > 0 ? (s.total / avgInq) * 50 : 50;
        const trialRateScore = avgTrial > 0 ? (sTrialRate / avgTrial) * 50 : 50;
        const enrRateScore = avgEnrRate > 0 ? (s.trialEnrollRate / avgEnrRate) * 50 : 50;
        const enrCountScore = avgEnrCount > 0 ? (s.enroll / avgEnrCount) * 50 : 50;
        
        s.totalScore = (inqScore + trialRateScore + enrRateScore + enrCountScore) / 4;
        
        if (s.totalScore >= 65) {
            s.rank = 'S'; s.rankClass = 'rank-s';
        } else if (s.totalScore >= 52) {
            s.rank = 'A'; s.rankClass = 'rank-a';
        } else if (s.totalScore >= 40) {
            s.rank = 'B'; s.rankClass = 'rank-b';
        } else {
            s.rank = 'C'; s.rankClass = 'rank-c';
        }
        
        // 苦手とする学年の特定 (入会率が教室の平均以下のものをすべて抽出)
        const avgEnrollRate = closedCount > 0 ? (s.enroll / closedCount) : 0;
        let strugglingGrades = [];
        
        // データが存在する学年の数をカウント
        const validGradesCount = Object.values(s.grades).filter(counts => (counts.enroll + counts.ng) > 0).length;
        
        // 比較対象の学年が2つ以上ある場合のみ苦手学年を判定
        if (validGradesCount > 1) {
            for (const [grade, counts] of Object.entries(s.grades)) {
                const gClosed = counts.enroll + counts.ng;
                if (gClosed > 0) {
                    const gradeEnrollRate = counts.enroll / gClosed;
                    // 教室の平均入会率以下の場合
                    if (gradeEnrollRate <= avgEnrollRate) {
                        strugglingGrades.push({
                            grade: grade,
                            rate: gradeEnrollRate,
                            enroll: counts.enroll,
                            trialDone: counts.trialDone // これは体験総数だが、A方式だとgClosedが本来の比較対象。ただ表示用にそのまま維持。
                        });
                    }
                }
            }
        }
        
        if (strugglingGrades.length === 0 || closedCount === 0) {
            s.worstGradeDisplay = '<span style="color:var(--text-muted);">なし</span>';
        } else {
            // 入会率が低い順にソート（同率なら体験数が多い順）
            strugglingGrades.sort((a, b) => {
                if (a.rate !== b.rate) return a.rate - b.rate;
                return b.trialDone - a.trialDone;
            });
            
            s.worstGradeDisplay = strugglingGrades.map(g => {
                const ratePct = (g.rate * 100).toFixed(0);
                return `<div style="margin-bottom:4px;"><span style="color:var(--danger); font-weight:bold;">${g.grade}</span> <span style="font-size:0.8rem; color:var(--text-secondary);">${ratePct}% (${g.enroll}/${g.trialDone}件)</span></div>`;
            }).join('');
        }
        
        return s;
    });

    // 指定されたキーと順序でソート
    summaryList.sort((a, b) => {
        let valA = a[storeTableSort.key];
        let valB = b[storeTableSort.key];
        
        // ランク列でソートする場合はスコア値で判定
        if (storeTableSort.key === 'rank') {
            valA = a.totalScore;
            valB = b.totalScore;
        }

        if (storeTableSort.key === 'name' || storeTableSort.key === 'courseType') {
            return storeTableSort.direction === 'asc' 
                ? valA.localeCompare(valB, 'ja') 
                : valB.localeCompare(valA, 'ja');
        }
        
        return storeTableSort.direction === 'asc' 
            ? valA - valB 
            : valB - valA;
    });

    // Excelエクスポート用にグローバル変数へ保存
    currentStoreSummaryList = summaryList;

    summaryList.forEach(s => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${s.name}</strong></td>
            <td>${s.courseType}</td>
            <td><span class="rank-badge ${s.rankClass}">${s.rank}</span></td>
            <td>${s.total}</td>
            <td>${s.trialDone}</td>
            <td><strong>${s.trialEnrollRate}%</strong></td>
            <td>${s.trialNgRate}%</td>
            <td><span class="badge badge-success">${s.enroll}</span></td>
            <td>${s.totalEnrollRate}%</td>
            <td>${s.ng} <span style="font-size:0.8rem; color:var(--text-secondary);">(${s.totalNgRate}%)</span></td>
            <td>${s.withdraw}</td>
        `;
        tbody.appendChild(tr);
    });

    if (summaryList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" style="text-align:center; color:var(--text-muted);">データがありません</td></tr>';
    }
}

// 生徒データ詳細テーブル (表示は最大200件に制限しパフォーマンスを維持)
function renderDetailTable(data) {
    const tbody = document.querySelector('#detail-table tbody');
    tbody.innerHTML = '';

    // ソート処理
    const sortedData = [...data].sort((a, b) => {
        const key = detailTableSort.key;
        let valA = a[key] !== undefined && a[key] !== null ? a[key] : '';
        let valB = b[key] !== undefined && b[key] !== null ? b[key] : '';

        // 日付文字列の場合は空文字をどう扱うかなどの処理が可能だが、今回は単純な文字列・数値比較
        if (typeof valA === 'string' && typeof valB === 'string') {
            valA = valA.trim();
            valB = valB.trim();
        }
        // idのソートは数値として
        if (key === 'id') {
            valA = Number(valA);
            valB = Number(valB);
        }

        if (valA < valB) return detailTableSort.direction === 'asc' ? -1 : 1;
        if (valA > valB) return detailTableSort.direction === 'asc' ? 1 : -1;
        return 0;
    });

    const maxDisplay = 200;
    const displayData = sortedData.slice(0, maxDisplay);

    displayData.forEach(item => {
        const tr = document.createElement('tr');
        
        let statusBadgeClass = 'badge-gray';
        if (item.status === '受講生') statusBadgeClass = 'badge-success';
        else if (item.status === 'NG') statusBadgeClass = 'badge-danger';
        else if (item.status === '退会生') statusBadgeClass = 'badge-warning';
        else if (item.status === '問合せ' || item.status === '体験') statusBadgeClass = 'badge-info';

        tr.innerHTML = `
            <td>${item.id}</td>
            <td><strong>${item.displayName}</strong></td>
            <td>${item.store}</td>
            <td><strong>${item.trialGrade}</strong></td>
            <td>${item.currentGrade || '(空白)'}</td>
            <td>${item.purpose}</td>
            <td><span class="badge ${statusBadgeClass}">${item.status}</span></td>
            <td>${item.inquiryDateStr || '-'}</td>
            <td>${item.trialDateStr || '-'}</td>
            <td>${item.enrollDateStr || '-'}</td>
            <td>${item.studyTime !== '' && item.studyTime !== undefined ? item.studyTime : '-'}</td>
        `;
        tbody.appendChild(tr);
    });

    if (data.length > maxDisplay) {
        const trInfo = document.createElement('tr');
        trInfo.innerHTML = `<td colspan="11" style="text-align:center; color:var(--text-muted); font-style:italic;">表示件数を最大 ${maxDisplay} 件に制限しています (全体 ${data.length} 件中)</td>`;
        tbody.appendChild(trInfo);
    }

    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" style="text-align:center; color:var(--text-muted);">データがありません</td></tr>';
    }
}

// ==========================================================================
// 7. 追加要件（エクスポート機能・AIチャット機能）
// ==========================================================================

// Excelへのエクスポート
function exportStoreTableToExcel() {
    if (currentStoreSummaryList.length === 0) {
        alert('出力するデータがありません。');
        return;
    }
    
    // エクスポート用データ配列の作成
    const exportRows = currentStoreSummaryList.map((s, idx) => {
        // 苦手とする学年のテキスト抽出
        const closedCount = s.enroll + s.ng;
        const avgEnrollRate = closedCount > 0 ? (s.enroll / closedCount) : 0;
        let strugglingGrades = [];
        const validGradesCount = Object.values(s.grades).filter(counts => (counts.enroll + counts.ng) > 0).length;
        
        if (validGradesCount > 1) {
            for (const [grade, counts] of Object.entries(s.grades)) {
                const gClosed = counts.enroll + counts.ng;
                if (gClosed > 0) {
                    const gradeEnrollRate = counts.enroll / gClosed;
                    if (gradeEnrollRate <= avgEnrollRate) {
                        strugglingGrades.push({
                            grade: grade,
                            rate: gradeEnrollRate,
                            enroll: counts.enroll,
                            trialDone: counts.trialDone
                        });
                    }
                }
            }
        }

        let worstGradeText = 'なし';
        if (strugglingGrades.length > 0) {
            strugglingGrades.sort((a, b) => {
                if (a.rate !== b.rate) return a.rate - b.rate;
                return b.trialDone - a.trialDone;
            });
            worstGradeText = strugglingGrades.map(g => {
                const ratePct = (g.rate * 100).toFixed(0);
                return `${g.grade}: ${ratePct}% (${g.enroll}/${g.trialDone}件)`;
            }).join(', ');
        }

        return {
            'No': idx + 1,
            '教室名': s.name,
            '受講種別': s.courseType,
            '総合ランク': s.rank,
            '総問合せ数': s.total,
            '体験実施数': s.trialDone,
            '体験入会率': `${s.trialEnrollRate}%`,
            '体験NG率': `${s.trialNgRate}%`,
            '入会数': s.enroll,
            '問合せ入会率': `${s.totalEnrollRate}%`,
            'NG数': s.ng,
            '退会者数': s.withdraw,
            '苦手とする学年 (平均入会率以下)': worstGradeText
        };
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "教室別集計");

    const baseDateInput = document.getElementById('base-date');
    const dateStr = baseDateInput && baseDateInput.value ? baseDateInput.value : new Date().toISOString().split('T')[0];
    const fileName = `教室別集計_${dateStr}.xlsx`;

    XLSX.writeFile(wb, fileName);
}

// 生徒詳細のExcelエクスポート
function exportDetailTableToExcel() {
    if (filteredData.length === 0) {
        alert('出力するデータがありません。');
        return;
    }
    
    // エクスポート用データ配列の作成
    const exportRows = filteredData.map((item, idx) => ({
        'No': idx + 1,
        '受講生ID': item.studentId || '',
        '氏名': item.displayName || '',
        '教室名': item.store || '',
        '体験時学年': item.trialGrade || '',
        '現在の学年': item.currentGrade || '',
        '来校目的': item.purpose || '',
        'ステータス (区分)': item.status || '',
        '問合せ日': item.inquiryDateStr || '',
        '体験日時': item.trialDateStr || '',
        '入会日': item.enrollDateStr || '',
        '受講時間': item.studyTime !== '' && item.studyTime !== undefined ? item.studyTime : ''
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "生徒データ詳細");

    const baseDateInput = document.getElementById('base-date');
    const dateStr = baseDateInput && baseDateInput.value ? baseDateInput.value : new Date().toISOString().split('T')[0];
    const fileName = `生徒データ詳細_${dateStr}.xlsx`;

    XLSX.writeFile(wb, fileName);
}

// チャートデータ表の切り替えとExcel出力処理
function initChartTableToggles() {
    const toggleGradeBtn = document.getElementById('toggle-grade-table-btn');
    const gradeContainer = document.getElementById('grade-table-container');
    const exportGradeBtn = document.getElementById('export-grade-excel-btn');
    if (toggleGradeBtn && gradeContainer) {
        toggleGradeBtn.addEventListener('click', () => {
            const isHidden = gradeContainer.style.display === 'none';
            gradeContainer.style.display = isHidden ? 'block' : 'none';
            toggleGradeBtn.textContent = isHidden ? '📊 表を非表示' : '📊 表を表示';
        });
    }

    const togglePurposeBtn = document.getElementById('toggle-purpose-table-btn');
    const purposeContainer = document.getElementById('purpose-table-container');
    const exportPurposeBtn = document.getElementById('export-purpose-excel-btn');
    if (togglePurposeBtn && purposeContainer) {
        togglePurposeBtn.addEventListener('click', () => {
            const isHidden = purposeContainer.style.display === 'none';
            purposeContainer.style.display = isHidden ? 'block' : 'none';
            togglePurposeBtn.textContent = isHidden ? '📊 表を非表示' : '📊 表を表示';
        });
    }

    const toggleYoyBtn = document.getElementById('toggle-yoy-table-btn');
    const toggleYoyBtn2 = document.getElementById('toggle-yoy-table-btn-2');
    const yoyContainer = document.getElementById('yoy-table-container');
    const exportYoyBtn = document.getElementById('export-yoy-excel-btn');
    
    const handleYoyToggle = () => {
        const isHidden = yoyContainer.style.display === 'none';
        yoyContainer.style.display = isHidden ? 'block' : 'none';
        const text = isHidden ? '📊 比較表を非表示' : '📊 比較表を表示';
        if (toggleYoyBtn) toggleYoyBtn.textContent = text;
        if (toggleYoyBtn2) toggleYoyBtn2.textContent = text;
    };

    if (toggleYoyBtn && yoyContainer) {
        toggleYoyBtn.addEventListener('click', handleYoyToggle);
    }
    if (toggleYoyBtn2 && yoyContainer) {
        toggleYoyBtn2.addEventListener('click', handleYoyToggle);
    }

    // YoY Inquiries / Enrolls mode toggle buttons
    const btnInq = document.getElementById('btn-yoy-mode-inq');
    const btnEnroll = document.getElementById('btn-yoy-mode-enroll');
    
    if (btnInq && btnEnroll) {
        btnInq.addEventListener('click', () => {
            if (yoyMode === 'inq') return;
            yoyMode = 'inq';
            btnInq.classList.add('active');
            btnEnroll.classList.remove('active');
            btnInq.style.background = 'var(--surface)';
            btnInq.style.color = 'var(--primary)';
            btnInq.style.boxShadow = 'var(--shadow-sm)';
            btnEnroll.style.background = 'transparent';
            btnEnroll.style.color = 'var(--text-secondary)';
            btnEnroll.style.boxShadow = 'none';
            // Redraw charts
            renderCharts(filteredData);
        });

        btnEnroll.addEventListener('click', () => {
            if (yoyMode === 'enroll') return;
            yoyMode = 'enroll';
            btnEnroll.classList.add('active');
            btnInq.classList.remove('active');
            btnEnroll.style.background = 'var(--surface)';
            btnEnroll.style.color = 'var(--primary)';
            btnEnroll.style.boxShadow = 'var(--shadow-sm)';
            btnInq.style.background = 'transparent';
            btnInq.style.color = 'var(--text-secondary)';
            btnInq.style.boxShadow = 'none';
            // Redraw charts
            renderCharts(filteredData);
        });
    }

    if (exportGradeBtn) {
        exportGradeBtn.addEventListener('click', () => {
            const table = document.getElementById('chart-grade-table');
            if (!table) return;
            const wb = XLSX.utils.table_to_book(table, {sheet: "学年別集計"});
            const baseDateInput = document.getElementById('base-date');
            const dateStr = baseDateInput && baseDateInput.value ? baseDateInput.value : new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `学年別集計_${dateStr}.xlsx`);
        });
    }

    if (exportPurposeBtn) {
        exportPurposeBtn.addEventListener('click', () => {
            const table = document.getElementById('chart-purpose-table');
            if (!table) return;
            const wb = XLSX.utils.table_to_book(table, {sheet: "来校目的別集計"});
            const baseDateInput = document.getElementById('base-date');
            const dateStr = baseDateInput && baseDateInput.value ? baseDateInput.value : new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `来校目的別集計_${dateStr}.xlsx`);
        });
    }

    if (exportYoyBtn) {
        exportYoyBtn.addEventListener('click', () => {
            const table = document.getElementById('chart-yoy-table');
            if (!table) return;
            const wb = XLSX.utils.table_to_book(table, {sheet: "前年比較集計"});
            const baseDateInput = document.getElementById('base-date');
            const dateStr = baseDateInput && baseDateInput.value ? baseDateInput.value : new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `前年比較集計_${dateStr}.xlsx`);
        });
    }
}

// APIキーの保存
function saveApiKey() {
    const keyInput = document.getElementById('gemini-api-key');
    const status = document.getElementById('api-key-status');
    const key = keyInput.value.trim();
    
    if (key) {
        localStorage.setItem('gemini_api_key', key);
        status.textContent = '設定済み';
        status.style.color = 'var(--success)';
    } else {
        localStorage.removeItem('gemini_api_key');
        status.textContent = '未設定';
        status.style.color = 'var(--text-muted)';
    }
}

// AIチャットへのメッセージ追加
function addChatMessage(message, sender) {
    const chatMessages = document.getElementById('chat-messages');
    const msgDiv = document.createElement('div');
    msgDiv.className = `chat-msg ${sender}`;
    
    // 簡易的なマークダウン（太字など）と改行のパース
    let htmlMessage = message.replace(/\\n/g, '<br>').replace(/\n/g, '<br>');
    htmlMessage = htmlMessage.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    
    msgDiv.innerHTML = htmlMessage;
    chatMessages.appendChild(msgDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// AIチャットの送信処理
async function handleAiChatSubmit() {
    const inputField = document.getElementById('chat-input-text');
    const question = inputField.value.trim();
    const apiKey = localStorage.getItem('gemini_api_key');
    
    if (!question) return;
    
    if (!apiKey) {
        alert("Gemini APIキーを入力して「保存」ボタンを押してください。");
        return;
    }
    
    if (filteredData.length === 0) {
        alert("分析対象のデータがありません。Excelを読み込んでいるか確認してください。");
        return;
    }

    // ユーザーのメッセージを表示
    addChatMessage(question, 'user');
    inputField.value = '';
    
    // 処理中メッセージ
    const loadingId = 'loading-' + Date.now();
    const chatMessages = document.getElementById('chat-messages');
    const loadingDiv = document.createElement('div');
    loadingDiv.className = `chat-msg ai`;
    loadingDiv.id = loadingId;
    loadingDiv.innerHTML = 'データ分析中... <span style="animation: blink 1s infinite;">⏳</span>';
    chatMessages.appendChild(loadingDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    // コンテキスト用データの作成
    // 教室別の集計情報をAIに渡す
    const storeMap = {};
    filteredData.forEach(item => {
        if (!storeMap[item.store]) {
            storeMap[item.store] = { 教室名: item.store, 総問合せ数: 0, 体験実施数: 0, 入会数: 0, NG数: 0, 退会数: 0 };
        }
        storeMap[item.store].総問合せ数++;
        const isNotTrial = ['問合せ', '体験キャンセル', '体験予約に至らず'].includes(item.status);
        if (!isNotTrial) storeMap[item.store].体験実施数++;
        if (item.status === '受講生' || item.status === '退会生') storeMap[item.store].入会数++;
        if (item.status === 'NG') storeMap[item.store].NG数++;
        if (item.status === '退会生') storeMap[item.store].退会数++;
    });
    
    let dataContext = "【教室別集計データ】\\n教室名,総問合せ数,体験実施数,入会数,体験入会率,NG数,退会数\\n";
    Object.values(storeMap).forEach(s => {
        const rate = s.体験実施数 > 0 ? (s.入会数 / s.体験実施数 * 100).toFixed(1) + "%" : "0.0%";
        dataContext += `${s.教室名},${s.総問合せ数},${s.体験実施数},${s.入会数},${rate},${s.NG数},${s.退会数}\\n`;
    });
    
    // もしフィルターが特定学年などで絞り込まれているならその情報も伝える
    const filterGrade = document.getElementById('filter-grade').value;
    const filterPurpose = document.getElementById('filter-purpose').value;
    const filterInfo = `現在このデータは以下の条件で絞り込まれた結果です。\\n学年: ${filterGrade}\\n目的: ${filterPurpose}\\n表示データ件数: ${filteredData.length}件`;
    
    try {
        const prompt = `あなたは優秀なデータアナリストです。以下の学習塾の教室別入会状況データを元に、ユーザーの質問に日本語で簡潔に回答してください。\\n\\n${filterInfo}\\n\\n${dataContext}\\n\\nユーザーの質問: ${question}`;
        
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }]
            })
        });
        
        const result = await response.json();
        
        const loadingElement = document.getElementById(loadingId);
        if (loadingElement) loadingElement.remove();
        
        if (!response.ok) {
            throw new Error(result.error?.message || 'API通信エラーが発生しました');
        }
        
        const aiMessage = result.candidates[0].content.parts[0].text;
        addChatMessage(aiMessage, 'ai');
        
    } catch (error) {
        console.error(error);
        const loadingElement = document.getElementById(loadingId);
        if (loadingElement) loadingElement.remove();
        addChatMessage(`エラーが発生しました: ${error.message}`, 'ai');
    }
}

// 印刷直前にChart.jsのサイズを印刷用レイアウト（160px）に強制リサイズさせるイベントリスナー
window.addEventListener('beforeprint', () => {
    for (const key in charts) {
        if (charts[key]) {
            charts[key].resize();
        }
    }
});
window.addEventListener('afterprint', () => {
    for (const key in charts) {
        if (charts[key]) {
            charts[key].resize();
        }
    }
});

// 現場ヒアリングガイドの表示トグル
function toggleHearingGuide(btn) {
    const box = btn.nextElementSibling;
    if (box) {
        const isHidden = box.style.display === 'none';
        box.style.display = isHidden ? 'block' : 'none';
        btn.textContent = isHidden ? '❌ ガイドを閉じる' : '📋 現場ヒアリング項目を確認';
    }
}
