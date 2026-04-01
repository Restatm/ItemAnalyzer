// Data State
let globalData = {
    items: [], // Array of item objects
    companies: [],
    categories: [],
    cvScores: {} // CV by category
};

let chartInstance = null;
let distChartInstance = null;
let currentFilters = { 
    high: true, 
    low: true, 
    polar: true,
    columnCat: [], // 빈 배열은 '전체'를 의미
    columnType: [],
    columnScore: [],
    columnReason: []
};

let activeFilterCol = null; // 현재 열려있는 필터 컬럼 키
let tempSelectedValues = new Set(); // 확인 버튼 누르기 전 임시 선택값
let currentViewMode = 'ALL'; // 'ALL', 'FLAGGED', or 'VIZ'
let selectedItemForSim = null;
let simulatedItemState = null; // Holds the state of the item being simulated

// 3D Visualization State
let vizState = {
    canvas: null,
    ctx: null,
    pts: [],
    categories: [],
    catColors: {
        "노동관행": "#5B7FDB",
        "직장 내 안전보건": "#3BAD8A",
        "인권": "#D95C5C",
        "공정운영관행": "#E5913A",
        "지속가능한 소비": "#9B6DD4",
        "정보보호 및 개인정보보호": "#D45C91",
        "지역사회 참여 및 개발": "#5BAAC4",
        "이해관계자 소통": "#A0A050"
    },
    activeCats: new Set(),
    rotX: -0.55,
    rotY: 3.92,
    zoom: 0.9,
    dragOn: false,
    lastMx: 0,
    lastMy: 0,
    hoveredPt: null,
    typeMap: { "정책": -1, "목표": 0, "위험관리": 1, "성과": 2 },
    RATE_SCALE: 8.0
};

// ... (DOM Elements and Event Listeners) ...
// We will replace the colors in the palettes object further down.

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const loading = document.getElementById('loading');
const dashboard = document.getElementById('app-dashboard');
const tableBody = document.getElementById('table-body');
const uploadSection = document.getElementById('upload-section');
const viewList = document.getElementById('view-list');
const viewSim = document.getElementById('view-sim');
const selectedItemName = document.getElementById('selected-item-id');
const scoreBtns = document.querySelectorAll('.score-btn');

// Event Listeners
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.style.borderColor = 'var(--primary-hover)'; dropZone.style.backgroundColor = 'rgba(79, 70, 229, 0.05)'; });
dropZone.addEventListener('dragleave', () => { dropZone.style.borderColor = ''; dropZone.style.backgroundColor = ''; });
dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.style.borderColor = ''; dropZone.style.backgroundColor = ''; if (e.dataTransfer.files.length > 0) processFile(e.dataTransfer.files[0]); });
fileInput.addEventListener('change', (e) => { if (e.target.files.length > 0) processFile(e.target.files[0]); });

document.getElementById('filter-high').addEventListener('change', (e) => { currentFilters.high = e.target.checked; renderTable(); });
document.getElementById('filter-low').addEventListener('change', (e) => { currentFilters.low = e.target.checked; renderTable(); });
document.getElementById('filter-polar').addEventListener('change', (e) => { currentFilters.polar = e.target.checked; renderTable(); });

// 신규 컬럼 필터 이벤트 리스너
// 다중 선택 필터 트리거 이벤트 바인딩
document.querySelectorAll('.filter-select').forEach(el => {
    el.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = el.getAttribute('data-col');
        openFilterMenu(colKey, el);
    });
});

// 외부 클릭 시 필터 메뉴 닫기
document.addEventListener('click', (e) => {
    const menu = document.getElementById('filter-menu');
    if (!menu.contains(e.target) && !e.target.classList.contains('filter-select')) {
        closeFilterMenu();
    }
});

// 필터 검색 이벤트
document.getElementById('filter-search').addEventListener('input', (e) => {
    renderFilterList(e.target.value);
});

// 스크롤 시 필터 메뉴 위치 실시간 업데이트 (Sticky 헤더 대응)
window.addEventListener('scroll', () => {
    updateFilterMenuPosition();
}, { capture: true, passive: true });

function updateFilterMenuPosition() {
    if (!activeFilterCol) return;
    const menu = document.getElementById('filter-menu');
    
    const triggerEl = document.querySelector(`.filter-select[data-col="${activeFilterCol}"]`);
    if (triggerEl) {
        const rect = triggerEl.getBoundingClientRect();
        menu.style.top = `${rect.bottom + 5}px`;
        menu.style.left = `${rect.left}px`;
    }
}

// 전체 선택 체크박스
document.getElementById('filter-select-all').addEventListener('change', (e) => {
    const isChecked = e.target.checked;
    const items = getUniqueValues(activeFilterCol);
    if (isChecked) {
        items.forEach(v => tempSelectedValues.add(String(v)));
    } else {
        tempSelectedValues.clear();
    }
    renderFilterList(document.getElementById('filter-search').value);
});

// 확인 버튼
document.getElementById('filter-apply-btn').addEventListener('click', () => {
    applyMultiFilter();
});

// 필터 초기화 버튼
document.getElementById('btn-reset-filters').addEventListener('click', () => {
    currentFilters.difficulty = 'all';
    currentFilters.polar = true;
    currentFilters.columnCat = [];
    currentFilters.columnType = [];
    currentFilters.columnScore = [];
    currentFilters.columnReason = [];
    
    // UI 초기화
    document.querySelectorAll('.filter-check').forEach(cb => {
        if (cb.value === 'all') cb.checked = true;
        else if (cb.id === 'check-polar') cb.checked = true;
        else cb.checked = false;
    });
    
    renderTable();
});

document.getElementById('card-total-items').addEventListener('click', () => {
    currentViewMode = 'ALL';
    
    // 시뮬레이션이나 시각화 화면이 열려있다면 닫기
    viewSim.classList.add('hidden');
    document.getElementById('view-viz').classList.add('hidden');
    viewList.classList.remove('hidden');
    
    // 카드 스타일 업데이트
    document.getElementById('card-total-items').style.borderColor = 'var(--primary)';
    document.getElementById('card-total-items').style.borderWidth = '2px';
    document.getElementById('card-flagged-items').style.borderColor = 'var(--border)';
    document.getElementById('card-flagged-items').style.borderWidth = '1px';
    document.getElementById('card-viz-items').style.borderColor = 'var(--border)';
    document.getElementById('card-viz-items').style.borderWidth = '1px';
    
    document.getElementById('table-title').textContent = '문항 전체';
    renderTable();
});
document.getElementById('card-flagged-items').addEventListener('click', () => {
    currentViewMode = 'FLAGGED';

    // 시뮬레이션이나 시각화 화면이 열려있다면 닫기
    viewSim.classList.add('hidden');
    document.getElementById('view-viz').classList.add('hidden');
    viewList.classList.remove('hidden');

    // 카드 스타일 업데이트
    document.getElementById('card-flagged-items').style.borderColor = 'var(--primary)';
    document.getElementById('card-flagged-items').style.borderWidth = '2px';
    document.getElementById('card-total-items').style.borderColor = 'var(--border)';
    document.getElementById('card-total-items').style.borderWidth = '1px';
    document.getElementById('card-viz-items').style.borderColor = 'var(--border)';
    document.getElementById('card-viz-items').style.borderWidth = '1px';
    
    document.getElementById('table-title').textContent = '검토 대상 문항';
    renderTable();
});
document.getElementById('card-viz-items').addEventListener('click', () => {
    openViz();
});

scoreBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
        scoreBtns.forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        updateSimulation();
    });
});

// TOP Button Logic
const btnTop = document.getElementById('btn-top');
window.addEventListener('scroll', () => {
    if (window.scrollY > 300) {
        btnTop.style.display = 'flex';
    } else {
        btnTop.style.display = 'none';
    }
});
btnTop.addEventListener('click', () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
});

// File Processing
function processFile(file) {
    loading.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            parseData(json);

            uploadSection.classList.add('hidden');
            dashboard.classList.remove('hidden');
            renderDashboard();
        } catch (error) {
            console.error(error);
            alert('엑셀 파일 파싱 중 오류가 발생했습니다. 데이터 구조를 확인하세요.');
        } finally {
            loading.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Logic to Parse Standardized Data Layout
function parseData(rows) {
    if (rows.length < 6) return; // Needs at least the 5 header rows + 1 data row

    const headers_cat = rows[0]; // 대분류
    const headers_id = rows[1];  // 번호
    const headers_type = rows[2]; // 유형
    const headers_max = rows[3]; // 만점
    const headers_q = rows[4];   // 문항

    let categories = [];
    let currentCategory = "기본분류";
    let itemCols = [];

    // Identify Item Columns
    for (let i = 0; i < headers_id.length; i++) {
        if (headers_cat[i] && headers_cat[i] !== "nan" && String(headers_cat[i]).trim() !== "") {
            currentCategory = headers_cat[i].toString().replace(/\.\d+$/, '').trim();
        }

        let cellId = headers_id[i];
        if (cellId && !isNaN(parseInt(cellId)) && parseInt(cellId) > 1000) {
            let maxS = parseFloat(headers_max[i]);
            itemCols.push({
                index: i,
                id: cellId.toString(),
                category: currentCategory,
                type: headers_type[i] ? headers_type[i].toString().trim() : '기본유형',
                maxScoreOrig: isNaN(maxS) ? 5 : maxS,
                question: headers_q[i] ? headers_q[i].toString().trim() : '',
            });
            if (!categories.includes(currentCategory)) categories.push(currentCategory);
        }
    }

    // Parse Company Responses (Handle '-' values as NULL)
    let companies = [];
    let totalRawCompaniesCount = 0;
    let excludedHoldingsCount = 0;

    // Find '지주사' column index dynamically to resist any Excel layout shifts
    let holdingTypeColIndex = -1;
    for (let i = 0; i < headers_cat.length; i++) {
        let hh = headers_cat[i] ? headers_cat[i].toString().trim() : "";
        if (hh === "지주사" || hh === "지주사 유형") {
            holdingTypeColIndex = i;
            break;
        }
    }

    for (let r = 5; r < rows.length; r++) { // Row index 5 starts company data
        if (!rows[r] || rows[r].length === 0) continue;
        if (!rows[r][0] || rows[r][0] === "nan" || String(rows[r][0]).trim() === "") continue;

        totalRawCompaniesCount++;

        // Exclude holding company types A and C
        if (holdingTypeColIndex !== -1) {
            let holdingType = rows[r][holdingTypeColIndex] ? String(rows[r][holdingTypeColIndex]).trim().toUpperCase() : "";
            if (holdingType === "A" || holdingType === "C") {
                excludedHoldingsCount++;
                continue;
            }
        }

        let compScores = {};
        itemCols.forEach(col => {
            let rawVal = rows[r][col.index];
            if (rawVal === undefined || rawVal === null || String(rawVal).trim() === '-' || String(rawVal).trim() === '') {
                compScores[col.id] = null; // Mark as explicit N/A (excluded from denominator)
            } else {
                let s = parseFloat(rawVal);
                compScores[col.id] = isNaN(s) ? null : s;
            }
        });
        companies.push({ name: rows[r][0], scores: compScores });
    }

    globalData.totalRawCompaniesCount = totalRawCompaniesCount;
    globalData.excludedHoldingsCount = excludedHoldingsCount;

    // Process Items and stats
    let items = itemCols.map(col => {
        // Only include companies that have valid scores for this specific item!!
        let scores = companies.map(c => c.scores[col.id]).filter(s => s !== null);
        let maxScore = col.maxScoreOrig;
        let actualMax = scores.length > 0 ? Math.max(...scores) : 0;
        // 엑셀 헤더의 만점과 실제 데이터의 최고점 중 큰 값을 만점으로 채택
        maxScore = Math.max(maxScore, actualMax);

        let sum = scores.reduce((a, b) => a + b, 0);
        let avg = scores.length > 0 ? sum / scores.length : 0;
        let scoringRate = maxScore > 0 ? (avg / maxScore) * 100 : 0;

        let distribution = {};
        scores.forEach(s => { distribution[s] = (distribution[s] || 0) + 1; });
        // 0점과 만점은 실제 득점 기업이 없더라도 항상 선지로 표시
        if (distribution[0] === undefined) distribution[0] = 0;
        if (maxScore > 0 && distribution[maxScore] === undefined) distribution[maxScore] = 0;
        let uniqueOptions = Object.keys(distribution).map(Number).sort((a, b) => a - b);

        // Polarization Rule (Rule 2)
        let isPolarized = false;
        if (uniqueOptions.length > 2) {
            let minVal = uniqueOptions[0];
            let maxVal = uniqueOptions[uniqueOptions.length - 1];
            let midOptions = uniqueOptions.filter(x => x !== minVal && x !== maxVal);

            let minSelectCount = Math.min(...Object.values(distribution));
            for (let mid of midOptions) {
                if (distribution[mid] === minSelectCount) {
                    isPolarized = true;
                }
            }
        }

        return {
            id: col.id,
            category: col.category,
            type: col.type,
            maxScore: maxScore,
            question: col.question,
            avg: avg,
            scoringRate: scoringRate,
            options: uniqueOptions,
            origOptions: [...uniqueOptions],
            distribution: distribution, // Map of score -> count
            isPolarized: isPolarized,
            flagReason: [],
            isHighestType: false,
            isLowestType: false
        };
    });

    // Rule 1: Difficulty Flagging (2-step logic)
    
    // Pass 1: Categorized Local Grouping (Category + Type)
    let typeGroups = {};
    items.forEach(it => {
        let groupKey = `${it.category}::${it.type}`;
        if (!typeGroups[groupKey]) typeGroups[groupKey] = [];
        typeGroups[groupKey].push(it);
    });

    for (const [key, typItems] of Object.entries(typeGroups)) {
        if (typItems.length <= 2) continue; 
        typItems.sort((a, b) => b.scoringRate - a.scoringRate);
        
        if (typItems[0].scoringRate >= 50) {
            typItems[0].isHighestType = true;
            typItems[0].flagReason.push(`[${typItems[0].type}] 난이도 하`);
        }

        let lastIdx = typItems.length - 1;
        if (typItems[lastIdx].scoringRate < 50) {
            typItems[lastIdx].isLowestType = true;
            typItems[lastIdx].flagReason.push(`[${typItems[0].type}] 난이도 상`);
        }
    }

    // Pass 2: Global Type Grouping (Across all categories for sparse types)
    let globalTypeGroups = {};
    items.forEach(it => {
        if (!globalTypeGroups[it.type]) globalTypeGroups[it.type] = [];
        globalTypeGroups[it.type].push(it);
    });

    for (const [typeName, gItems] of Object.entries(globalTypeGroups)) {
        if (gItems.length <= 1) continue;
        gItems.sort((a, b) => b.scoringRate - a.scoringRate);

        // Global Highest in this Type (if not already flagged)
        if (gItems[0].scoringRate >= 50 && !gItems[0].isHighestType) {
            gItems[0].isHighestType = true;
            gItems[0].flagReason.push(`[${typeName}] 난이도 하`);
        }
        // Global Lowest in this Type (if not already flagged)
        let lastIdx = gItems.length - 1;
        if (gItems[lastIdx].scoringRate < 50 && !gItems[lastIdx].isLowestType) {
            gItems[lastIdx].isLowestType = true;
            gItems[lastIdx].flagReason.push(`[${typeName}] 난이도 상`);
        }
    }

    items.forEach(it => {
        if (it.isPolarized) it.flagReason.push('양극화');
    });

    globalData.items = items;
    globalData.companies = companies;
    globalData.categories = categories;

    globalData.cvScores = calculateCV(globalData.items);
}

// Target CV Calculation Function
function calculateCV(itemsSet) {
    let cvs = {};

    globalData.categories.forEach(cat => {
        let catItems = itemsSet.filter(it => it.category === cat);
        if (catItems.length === 0) {
            cvs[cat] = 0; return;
        }

        let totalMaxScore = catItems.reduce((acc, it) => acc + it.maxScore, 0);
        if (totalMaxScore === 0) {
            cvs[cat] = 0; return;
        }

        let typeSums = {};
        catItems.forEach(it => {
            typeSums[it.type] = (typeSums[it.type] || 0) + it.maxScore;
        });

        let proportions = [];
        for (let typ in typeSums) {
            proportions.push((typeSums[typ] / totalMaxScore) * 100);
        }

        if (proportions.length < 2) {
            cvs[cat] = 0; return;
        }

        let mean = proportions.reduce((a, b) => a + b, 0) / proportions.length;
        let sumSq = proportions.reduce((a, b) => a + Math.pow(b - mean, 2), 0);
        
        // Sample Standard Deviation (Divide by n-1)
        // If n <= 1, standard deviation is not defined (set to 0)
        let stdDev = proportions.length > 1 ? Math.sqrt(sumSq / (proportions.length - 1)) : 0;

        cvs[cat] = stdDev / mean;
    });

    return cvs;
}

function renderDashboard() {
    const items = globalData.items;
    const flagged = items.filter(it => it.flagReason.length > 0);
    const nonTrialCount = items.filter(it => it.maxScore > 0).length;
    
    // 상세 집계 (검토 대상 문항용)
    const polarCount = items.filter(it => it.isPolarized).length;
    const lowCount = items.filter(it => it.isLowestType).length;
    const highCount = items.filter(it => it.isHighestType).length;

    // 1. 전체 문항 카드
    document.getElementById('stat-total-items').innerHTML = `
        ${items.length} <span class="stat-badge">시범 제외<span class="badge-val">${nonTrialCount}</span></span>
    `;
    
    // 2. 검토 대상 문항 카드 (세로 스택 멀티 배지)
    document.getElementById('stat-flagged-items').innerHTML = `
        ${flagged.length} 
        <span class="stat-badge" style="background:#fff7ed; color:#c2410c;">양극화<span class="badge-val">${polarCount}</span></span>
        <span class="stat-badge" style="background:#f0fdf4; color:#15803d;">난이도 하<span class="badge-val">${lowCount}</span></span>
        <span class="stat-badge" style="background:#fef2f2; color:#b91c1c;">난이도 상<span class="badge-val">${highCount}</span></span>
    `;
    
    // 3. 참여 기업 수 카드
    document.getElementById('stat-total-companies').innerHTML = `
        ${globalData.totalRawCompaniesCount} 
        <span class="stat-badge">지주 제외<span class="badge-val">${globalData.companies.length}</span></span>
    `;

    renderTable();
    populateColumnFilters(items);
    renderCVChart(globalData.cvScores);
}

// 테이블 헤더 필터 드롭다운 옵션 생성
function populateColumnFilters() {
    // 다중 선택 필터로 개편되어 더 이상 select 박스를 여기서 채우지 않음
    renderTable(); 
}

function getUniqueValues(colKey) {
    if (!globalData.items) return [];
    if (colKey === 'columnCat') return [...new Set(globalData.items.map(it => it.category))].sort();
    if (colKey === 'columnType') return [...new Set(globalData.items.map(it => it.type))].sort();
    if (colKey === 'columnScore') return [...new Set(globalData.items.map(it => it.maxScore))].sort((a,b) => a-b).map(String);
    if (colKey === 'columnReason') return [...new Set(globalData.items.flatMap(it => it.flagReason))].sort();
    return [];
}

function openFilterMenu(colKey, triggerEl) {
    activeFilterCol = colKey;
    const menu = document.getElementById('filter-menu');
    const searchInput = document.getElementById('filter-search');
    searchInput.value = '';
    
    // 현재 선택된 값들을 세트에 복사 (임시 편집용)
    if (currentFilters[colKey].length === 0) {
        // 필터가 비어있다면(전체보기 상태) 모든 유니크 값을 체크된 상태로 시작
        tempSelectedValues = new Set(getUniqueValues(colKey).map(String));
    } else {
        tempSelectedValues = new Set(currentFilters[colKey]);
    }
    
    // 위치 조정 전 메뉴 노출 (위치 계산 함수에서 'hidden' 상태면 중단되는 로직 대비)
    menu.classList.remove('hidden');
    updateFilterMenuPosition();
    
    renderFilterList();
}

function closeFilterMenu() {
    document.getElementById('filter-menu').classList.add('hidden');
    activeFilterCol = null;
}

function renderFilterList(searchText = '') {
    const listCont = document.getElementById('filter-list');
    listCont.innerHTML = '';
    const items = getUniqueValues(activeFilterCol);
    const filtered = items.filter(v => String(v).toLowerCase().includes(searchText.toLowerCase()));
    
    // 전체 선택 체크박스 상태 업데이트
    const allSelected = filtered.length > 0 && filtered.every(v => tempSelectedValues.has(String(v)));
    document.getElementById('filter-select-all').checked = allSelected;

    filtered.forEach(val => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'filter-item';
        const isChecked = tempSelectedValues.has(String(val));
        
        itemDiv.innerHTML = `
            <input type="checkbox" id="chk-${val}" ${isChecked ? 'checked' : ''}>
            <label for="chk-${val}" style="flex:1; cursor:pointer;">${val}</label>
        `;
        
        const chk = itemDiv.querySelector('input');
        chk.addEventListener('change', () => {
            if (chk.checked) tempSelectedValues.add(String(val));
            else tempSelectedValues.delete(String(val));
            
            // 실시간 '전체 선택' 상태 갱신
            const currentFiltered = getUniqueValues(activeFilterCol).filter(v => String(v).toLowerCase().includes(document.getElementById('filter-search').value.toLowerCase()));
            document.getElementById('filter-select-all').checked = currentFiltered.every(v => tempSelectedValues.has(String(v)));
        });

        itemDiv.addEventListener('click', (e) => {
            if (e.target !== chk) {
                chk.checked = !chk.checked;
                chk.dispatchEvent(new Event('change'));
            }
        });

        listCont.appendChild(itemDiv);
    });
}

function applyMultiFilter() {
    if (!activeFilterCol) return;
    currentFilters[activeFilterCol] = [...tempSelectedValues];
    closeFilterMenu();
    renderTable();
}

function renderTable() {
    tableBody.innerHTML = '';
    
    // 필터 활성화 상태에 따른 UI 강조 (보라색 그라데이션)
    const updateFilterUI = (id, activeList) => {
        const el = document.getElementById(id);
        if (el) {
            if (activeList && activeList.length > 0) el.classList.add('active-filter');
            else el.classList.remove('active-filter');
        }
    };
    updateFilterUI('filter-col-cat', currentFilters.columnCat);
    updateFilterUI('filter-col-type', currentFilters.columnType);
    updateFilterUI('filter-col-score', currentFilters.columnScore);
    updateFilterUI('filter-col-reason', currentFilters.columnReason);

    let displayItems = globalData.items.filter(it => {
        // 1. 대시보드 뷰 필터링
        if (currentViewMode === 'FLAGGED' && it.flagReason.length === 0) return false;

        // 2. 난이도/양극화 필터링 (OR 조건으로 하나라도 해당되면 통과)
        if (currentViewMode === 'FLAGGED') {
            let passHigh = currentFilters.high && it.isLowestType;
            let passLow = currentFilters.low && it.isHighestType;
            let passPolar = currentFilters.polar && it.isPolarized;
            if (!(passHigh || passLow || passPolar)) return false;
        }

        // 3. 다중 선택 컬럼 필터링 (비어있지 않은 경우에만 매칭 확인)
        // 엑셀과 동일한 동작을 위해 데이터 타입에 상관없이 String으로 변환하여 비교
        if (currentFilters.columnCat.length > 0 && !currentFilters.columnCat.includes(String(it.category))) return false;
        if (currentFilters.columnType.length > 0 && !currentFilters.columnType.includes(String(it.type))) return false;
        if (currentFilters.columnScore.length > 0 && !currentFilters.columnScore.includes(String(it.maxScore))) return false;
        
        // 4. 검토 사유 필터링 (교집합 확인)
        if (currentFilters.columnReason.length > 0) {
            const hasMatch = it.flagReason.some(r => currentFilters.columnReason.includes(String(r)));
            if (!hasMatch) return false;
        }

        return true;
    });

    displayItems.forEach(it => {
        let isTrial = it.maxScore === 0;
        let tr = document.createElement('tr');
        if (isTrial) tr.style.color = '#94a3b8';
        
        let polarBadge = it.isPolarized ? '<span class="badge warning">양극화</span>' : '-';
        let btnText = isTrial ? '시범문항입니다' : '분석 및 시뮬레이션';
        let btnClass = isTrial ? 'btn-trial' : 'btn-primary';
        let actionBtn = `<button class="btn btn-small btn-sim-action ${btnClass}" onclick="openSim('${it.id}')">${btnText}</button>`;
        let reasonStr = it.flagReason.length > 0 ? it.flagReason.join('<br>') : '<span style="color:#94a3b8">특이사항 없음</span>';

        tr.innerHTML = `
            <td>${it.category}</td>
            <td>${it.type}</td>
            <td title="${it.question}"><strong>${it.id}</strong></td>
            <td>${it.maxScore}점</td>
            <td>${it.scoringRate.toFixed(1)}%</td>
            <td>${polarBadge}</td>
            <td style="color:${isTrial ? '#cbd5e1' : '#d946ef'}; font-weight:${isTrial ? '400' : '700'}; font-size: 0.85rem">${reasonStr}</td>
            <td>${actionBtn}</td>
        `;
        tableBody.appendChild(tr);
    });
}

window.openSim = function (itemId) {
    let item = globalData.items.find(i => i.id === itemId);
    if (!item) return;

    selectedItemForSim = item;
    simulatedItemState = {
        maxScore: item.maxScore,
        optionsMap: {}
    };

    // Store valid company scores only
    simulatedItemState.origCompanyOptions = globalData.companies.map(c => c.scores[itemId]);
    simulatedItemState.activeOptions = [...item.origOptions];

    viewList.classList.add('hidden');
    viewSim.classList.remove('hidden');

    selectedItemName.textContent = `문항 ${item.id} (${item.type})`;

    // 배점 버튼을 문항의 만점에 맞게 동적 생성
    const scoreSelector = document.querySelector('#view-sim .score-selector');
    scoreSelector.innerHTML = '';
    let btnValues = [2, 3, 5, 7].filter(v => v <= item.maxScore);
    if (!btnValues.includes(item.maxScore)) btnValues.push(item.maxScore);
    btnValues.sort((a, b) => a - b);
    btnValues.forEach(val => {
        let btn = document.createElement('button');
        btn.className = 'score-btn';
        btn.setAttribute('data-val', val);
        btn.style.cssText = 'padding: 0.75rem 1.5rem; font-size: 1rem;';
        btn.textContent = val + '점';
        if (val === item.maxScore) btn.classList.add('active');
        btn.addEventListener('click', () => {
            scoreSelector.querySelectorAll('.score-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            updateSimulation();
        });
        scoreSelector.appendChild(btn);
    });

    updateSimulation();
    renderCVChart(globalData.cvScores); // Reset to base chart
}

window.closeSim = function () {
    viewSim.classList.add('hidden');
    viewList.classList.remove('hidden');
    selectedItemForSim = null;
    simulatedItemState = null;
    
    // Reset cards
    document.getElementById('card-total-items').style.borderColor = currentViewMode === 'ALL' ? 'var(--primary)' : 'var(--border)';
    document.getElementById('card-flagged-items').style.borderColor = currentViewMode === 'FLAGGED' ? 'var(--primary)' : 'var(--border)';
}

window.openViz = function() {
    currentViewMode = 'VIZ';
    viewList.classList.add('hidden');
    viewSim.classList.add('hidden');
    document.getElementById('view-viz').classList.remove('hidden');
    
    // Update card styles
    document.getElementById('card-viz-items').style.borderColor = 'var(--primary)';
    document.getElementById('card-viz-items').style.borderWidth = '2px';
    document.getElementById('card-total-items').style.borderColor = 'var(--border)';
    document.getElementById('card-total-items').style.borderWidth = '1px';
    document.getElementById('card-flagged-items').style.borderColor = 'var(--border)';
    document.getElementById('card-flagged-items').style.borderWidth = '1px';
    
    initViz();
}

window.closeViz = function() {
    currentViewMode = 'ALL';
    document.getElementById('view-viz').classList.add('hidden');
    viewList.classList.remove('hidden');
    
    // Reset cards
    document.getElementById('card-viz-items').style.borderColor = 'var(--border)';
    document.getElementById('card-viz-items').style.borderWidth = '1px';
    document.getElementById('card-total-items').style.borderColor = 'var(--primary)';
    document.getElementById('card-total-items').style.borderWidth = '2px';
}

window.removeOption = function (optVal) {
    if (!simulatedItemState) return;
    simulatedItemState.activeOptions = simulatedItemState.activeOptions.filter(o => o !== optVal);

    // Dynamic selector to find current active score buttons in the simulation studio
    let currentScoreButtons = document.querySelectorAll('#view-sim .score-btn');
    
    // If the removed option matches the currently active manual score button, unselect it.
    currentScoreButtons.forEach(btn => {
        if (btn.classList.contains('active') && parseInt(btn.getAttribute('data-val')) === optVal) {
            btn.classList.remove('active');
        }
    });

    updateSimulation();
}

window.restoreOption = function (optVal) {
    if (!simulatedItemState) return;
    if (!simulatedItemState.activeOptions.includes(optVal)) {
        simulatedItemState.activeOptions.push(optVal);
        simulatedItemState.activeOptions.sort((a, b) => a - b);
    }

    // Reset manual score button selection to let it follow the natural max when restoring
    let currentScoreButtons = document.querySelectorAll('#view-sim .score-btn');
    currentScoreButtons.forEach(btn => btn.classList.remove('active'));

    updateSimulation();
}

function updateSimulation() {
    if (!simulatedItemState) return;

    // 1. Identify Selected Max Score 
    const activeBtn = document.querySelector('#view-sim .score-btn.active');
    let currentActiveMax = Math.max(...simulatedItemState.activeOptions);
    // If a manual button is active, that overrides the natural max of remaining options.
    let newMaxScore = activeBtn ? parseInt(activeBtn.getAttribute('data-val')) : currentActiveMax;
    simulatedItemState.maxScore = newMaxScore;

    // 2. Map active options to new max score
    let scaledOptionsMap = {};
    simulatedItemState.activeOptions.forEach(opt => {
        if (opt === currentActiveMax) {
            scaledOptionsMap[opt] = newMaxScore; // Max score scales exactly to new Max
        } else if (opt > newMaxScore) {
            scaledOptionsMap[opt] = newMaxScore; // Intermediate options > new max cap out at max
        } else {
            scaledOptionsMap[opt] = opt; // Everything else stays fixed
        }
    });
    simulatedItemState.scaledOptionsMap = scaledOptionsMap;

    let scaledActiveOpts = simulatedItemState.activeOptions.map(opt => scaledOptionsMap[opt]);

    // 3. Process Option Fallbacks & Scaling for Scoring Rate
    let validCompanyOptions = simulatedItemState.origCompanyOptions.filter(v => v !== null);

    let newScores = validCompanyOptions.map(origVal => {
        let currentVal = origVal;
        // 1. apply fallback (deleted options cascading down)
        while (!simulatedItemState.activeOptions.includes(currentVal) && currentVal > 0) {
            let smallerOpts = simulatedItemState.activeOptions.filter(o => o < currentVal).sort((a, b) => b - a);
            currentVal = smallerOpts.length > 0 ? smallerOpts[0] : 0;
        }
        // 2. apply linear scale for MaxScore change
        return scaledOptionsMap[currentVal];
    });

    let newSum = newScores.reduce((a, b) => a + b, 0);
    let newAvg = newScores.length > 0 ? newSum / newScores.length : 0;

    let newScoringRate = newMaxScore > 0 ? (newAvg / newMaxScore) * 100 : 0;
    simulatedItemState.currentScoringRate = newScoringRate;

    // Distribution mapped to active options (scaled)
    let newDist = {};
    scaledActiveOpts.forEach(o => newDist[o] = 0);
    newScores.forEach(s => {
        if (newDist[s] !== undefined) newDist[s]++;
        else newDist[s] = 1;
    });
    simulatedItemState.currentDist = newDist;
    simulatedItemState.scaledActiveOpts = scaledActiveOpts;

    // Process CV recalculation (Since CV depends on Max Score)
    let clonedItems = globalData.items.map(it => {
        if (it.id === selectedItemForSim.id) {
            return { ...it, maxScore: newMaxScore };
        }
        return it;
    });

    let newCvs = calculateCV(clonedItems);

    updateSimulationUI();
    renderCVChart(globalData.cvScores, newCvs);
}

function updateSimulationUI() {
    let item = selectedItemForSim;
    if (!item || !simulatedItemState) return;

    let dist = simulatedItemState.currentDist;
    let srate = simulatedItemState.currentScoringRate;
    let activeOpts = simulatedItemState.activeOptions;
    let minOptsIdx = item.origOptions[0];

    const distContainer = document.getElementById('sim-distribution');
    distContainer.innerHTML = '';

    let validCompaniesCount = simulatedItemState.origCompanyOptions.filter(v => v !== null).length;

    item.origOptions.forEach(opt => {
        let isRemoved = !activeOpts.includes(opt);
        let scaledOpt = isRemoved ? opt : simulatedItemState.scaledOptionsMap[opt];
        let count = isRemoved ? 0 : (dist[scaledOpt] || 0);

        let perc = validCompaniesCount > 0 ? ((count / validCompaniesCount) * 100).toFixed(1) : 0;

        let card = document.createElement('div');
        card.className = `dist-card ${isRemoved ? 'removed' : ''}`;

        // Action button logic: Allow removing or restoring anything EXCEPT the absolute minimum
        let actionBtnHtml = '';
        if (opt !== minOptsIdx) {
            if (!isRemoved) {
                // Delete button
                actionBtnHtml = `<button class="btn btn-del" onclick="removeOption(${opt})" style="font-size:0.7rem; padding: 0.15rem 0.4rem; margin-top:0.4rem; background: #fef2f2; border: 1px solid #fca5a5; color: #ef4444; border-radius: 4px; cursor: pointer; width: 100%;">삭제</button>`;
            } else {
                // Restore button
                actionBtnHtml = `<button class="btn btn-restore" onclick="restoreOption(${opt})" style="font-size:0.7rem; padding: 0.15rem 0.4rem; margin-top:0.4rem; background: #e0f2fe; border: 1px solid #7dd3fc; color: #0284c7; border-radius: 4px; cursor: pointer; width: 100%;">복원</button>`;
            }
        }

        card.innerHTML = `
            <div style="font-size: 1rem; font-weight:700; color:var(--text-main)">${scaledOpt}점</div>
            <div style="font-size: 0.9rem; color:var(--primary); font-weight: 600; margin-top: 0.2rem;">${perc}%</div>
            <div style="font-size: 0.75rem; color:var(--text-muted)">(${count}사)</div>
            ${actionBtnHtml}
        `;
        distContainer.appendChild(card);
    });

    const srateContainer = document.getElementById('sim-scoring-rate');
    let origRate = item.scoringRate.toFixed(1);
    let newRate = srate.toFixed(1);

    if (origRate !== newRate) {
        srateContainer.innerHTML = `${origRate}% <span style="margin: 0 0.5rem">➔</span> <span class="highlight-red">${newRate}%</span>`;
    } else {
        srateContainer.innerHTML = `${origRate}% <span style="font-size: 0.85rem; color:#64748b; font-weight: 500; margin-left: 0.5rem;">(변동 없음)</span>`;
    }

    // Chart logic (handling duplicate merging and sorting)
    let chartDataMap = {};
    let uniqueChartOptions = [];

    item.origOptions.forEach(opt => {
        let isRemoved = !activeOpts.includes(opt);
        if (isRemoved) {
            uniqueChartOptions.push({ label: `${opt}점 (삭제됨)`, optVal: opt, isRemoved: true, count: 0, perc: "0.0" });
        } else {
            let scaledOpt = simulatedItemState.scaledOptionsMap[opt];
            if (!chartDataMap[scaledOpt]) {
                let count = dist[scaledOpt] || 0;
                let perc = validCompaniesCount > 0 ? ((count / validCompaniesCount) * 100).toFixed(1) : "0.0";
                chartDataMap[scaledOpt] = { label: `${scaledOpt}점`, optVal: scaledOpt, isRemoved: false, count: count, perc: perc };
            }
        }
    });

    Object.values(chartDataMap).forEach(info => {
        uniqueChartOptions.push(info);
    });
    // Sort by integer value
    uniqueChartOptions.sort((a, b) => a.optVal - b.optVal);

    let labels = uniqueChartOptions.map(u => u.label);
    let dataPercentages = uniqueChartOptions.map(u => u.perc);
    let removedFlags = uniqueChartOptions.map(u => u.isRemoved);
    let counts = uniqueChartOptions.map(u => u.count);

    renderDistChart(labels, dataPercentages, removedFlags, counts);
}

const customDataLabelsPlugin = {
    id: 'customDataLabels',
    afterDatasetsDraw(chart, args, pluginOptions) {
        const { ctx, data } = chart;
        ctx.save();
        ctx.font = 'bold 0.85rem Pretendard, sans-serif';
        ctx.fillStyle = '#475569';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        chart.getDatasetMeta(0).data.forEach((bar, index) => {
            let val = data.datasets[0].data[index];
            let count = data.datasets[0].counts[index];
            let isRemoved = data.datasets[0].removedFlags[index];
            if (!isRemoved) { // Only show label on active bars
                // 두 줄로 나누어 표기 (괄호 추가)
                ctx.fillText(`${count}사`, bar.x, bar.y - 20);
                ctx.fillText(`(${val}%)`, bar.x, bar.y - 4);
            }
        });
        ctx.restore();
    }
};

const palettes = {
    1: ['rgba(252, 165, 165, 0.8)'],
    2: ['rgba(252, 165, 165, 0.8)', 'rgba(147, 197, 253, 0.8)'], // Red, Blue
    3: ['rgba(252, 165, 165, 0.8)', 'rgba(254, 240, 138, 0.8)', 'rgba(134, 239, 172, 0.8)'], // Red, Yellow (Lighter), Green
    4: ['rgba(252, 165, 165, 0.8)', 'rgba(254, 240, 138, 0.8)', 'rgba(134, 239, 172, 0.8)', 'rgba(147, 197, 253, 0.8)'],
    5: ['rgba(252, 165, 165, 0.8)', 'rgba(253, 186, 116, 0.8)', 'rgba(254, 240, 138, 0.8)', 'rgba(134, 239, 172, 0.8)', 'rgba(147, 197, 253, 0.8)']
};

const borderPalettes = {
    1: ['#f87171'],
    2: ['#f87171', '#60a5fa'],
    3: ['#f87171', '#fde047', '#4ade80'],
    4: ['#f87171', '#fde047', '#4ade80', '#60a5fa'],
    5: ['#f87171', '#fb923c', '#fde047', '#4ade80', '#60a5fa']
};

function renderDistChart(labels, data, removedFlags, counts) {
    const ctx = document.getElementById('distChart').getContext('2d');

    let activeCount = removedFlags.filter(f => !f).length;
    let selectedPalette = palettes[activeCount] || palettes[5];
    let selectedBorders = borderPalettes[activeCount] || borderPalettes[5];

    let bgIndex = 0;
    let backgroundColors = removedFlags.map(isRemoved => {
        if (isRemoved) return 'rgba(203, 213, 225, 0.4)';
        let color = selectedPalette[bgIndex] || 'rgba(14, 165, 233, 0.6)';
        bgIndex++;
        return color;
    });

    let bdIndex = 0;
    let borderColors = removedFlags.map(isRemoved => {
        if (isRemoved) return '#cbd5e1';
        let color = selectedBorders[bdIndex] || '#0ea5e9';
        bdIndex++;
        return color;
    });

    // 동적 y축 max 계산
    let maxDataVal = Math.max(...data.map(Number));
    let yMax;
    if (maxDataVal <= 50) yMax = 60;
    else if (maxDataVal <= 70) yMax = 80;
    else yMax = 100;

    if (distChartInstance) {
        distChartInstance.data.labels = labels;
        distChartInstance.data.datasets[0].data = data;
        distChartInstance.data.datasets[0].counts = counts;
        distChartInstance.data.datasets[0].removedFlags = removedFlags;
        distChartInstance.data.datasets[0].backgroundColor = backgroundColors;
        distChartInstance.data.datasets[0].borderColor = borderColors;
        distChartInstance.options.scales.y.max = yMax;
        distChartInstance.update();
    } else {
        distChartInstance = new Chart(ctx, {
            type: 'bar',
            plugins: [customDataLabelsPlugin],
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    counts: counts,
                    removedFlags: removedFlags,
                    backgroundColor: backgroundColors,
                    borderColor: borderColors,
                    borderWidth: 1,
                    barPercentage: 0.5, // Reduces bar thickness further (~50%)
                    categoryPercentage: 0.8,
                    maxBarThickness: 60 // Prevents visually fat bars on wide screens
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 40 } }, // Adds space for the 2-line data labels
                scales: {
                    x: {
                        grid: { display: false }
                    },
                    y: {
                        beginAtZero: true,
                        max: yMax,
                        title: { display: false },
                        grid: { display: false, drawBorder: false },
                        border: { display: true, color: '#cbd5e1' },
                        ticks: {
                            display: true,
                            stepSize: 20,
                            callback: function (value) {
                                return value + '%';
                            }
                        }
                    }
                },
                plugins: {
                    legend: { display: false }
                },
                animation: {
                    duration: 500,
                    easing: 'easeOutQuart'
                }
            }
        });
    }
}

function renderCVChart(baseCvs, simCvs = null) {
    const ctx = document.getElementById('cvChart').getContext('2d');

    let labels = globalData.categories.map(c => c.length > 6 ? c.substring(0, 6) + '..' : c);
    let baseData = globalData.categories.map(c => baseCvs[c] !== undefined ? baseCvs[c].toFixed(2) : 0);

    if (chartInstance) {
        chartInstance.data.labels = labels;
        chartInstance.data.datasets[0].data = baseData; // base dataset

        if (simCvs) {
            let simData = globalData.categories.map(c => simCvs[c] !== undefined ? simCvs[c].toFixed(2) : 0);

            if (chartInstance.data.datasets.length === 1) {
                // Add secondary dataset
                chartInstance.data.datasets.push({
                    label: '개정 후 CV',
                    data: simData,
                    backgroundColor: 'rgba(79, 70, 229, 0.6)',
                    borderColor: '#4f46e5',
                    borderWidth: 2
                });
            } else {
                // Mutate existing secondary dataset for seamless transition animation
                chartInstance.data.datasets[1].data = simData;
            }
        } else {
            if (chartInstance.data.datasets.length === 2) {
                chartInstance.data.datasets.pop();
            }
        }
        chartInstance.update();
    } else {
        let datasets = [{
            label: '현재 CV (유형별 배점 변동계수)',
            data: baseData,
            backgroundColor: 'rgba(100, 116, 139, 0.4)',
            borderColor: '#64748b',
            borderWidth: 1
        }];

        if (simCvs) {
            let simData = globalData.categories.map(c => simCvs[c] !== undefined ? simCvs[c].toFixed(2) : 0);
            datasets.push({
                label: '개정 후 CV',
                data: simData,
                backgroundColor: 'rgba(79, 70, 229, 0.6)',
                borderColor: '#4f46e5',
                borderWidth: 2
            });
        }

        chartInstance = new Chart(ctx, {
            type: 'bar',
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        grid: { display: false }
                    },
                    y: {
                        beginAtZero: true,
                        title: { display: false },
                        grid: { display: false, drawBorder: false }
                    }
                },
                plugins: {
                    legend: { position: 'top' }
                },
                animation: {
                    duration: 500,
                    easing: 'easeOutQuart'
                }
            }
        });
    }
}

// ── 3D Visualization Implementation ──

function initViz() {
    if (!vizState.canvas) {
        vizState.canvas = document.getElementById("vizCanvas");
        vizState.ctx = vizState.canvas.getContext("2d");
        
        // Mouse Events
        vizState.canvas.addEventListener("mousedown", e => {
            vizState.dragOn = true; 
            vizState.lastMx = e.clientX; 
            vizState.lastMy = e.clientY;
        });
        window.addEventListener("mouseup", () => { vizState.dragOn = false; });
        window.addEventListener("mousemove", handleVizMouseMove);
        vizState.canvas.addEventListener("wheel", e => {
            e.preventDefault();
            // 휠 감도 및 범위 최적화 (0.2배 ~ 5.0배)
            const delta = e.deltaY > 0 ? 0.9 : 1.1; 
            vizState.zoom = Math.max(0.2, Math.min(5.0, vizState.zoom * delta));
            drawViz();
        }, { passive: false });
        
        // Touch Events
        let lastTX=0, lastTY=0;
        vizState.canvas.addEventListener("touchstart", e => {
            lastTX = e.touches[0].clientX;
            lastTY = e.touches[0].clientY;
        }, { passive: true });
        vizState.canvas.addEventListener("touchmove", e => {
            vizState.rotY += (e.touches[0].clientX - lastTX) * 0.01;
            vizState.rotX += (e.touches[0].clientY - lastTY) * 0.01;
            lastTX = e.touches[0].clientX;
            lastTY = e.touches[0].clientY;
            drawViz();
        }, { passive: true });
    }
    
    // Prepare Data (Filter out Trial Items: maxScore > 0)
    const items = globalData.items.filter(it => it.maxScore > 0);
    const typeMap = vizState.typeMap;
    const RATE_SCALE = vizState.RATE_SCALE;
    
    vizState.pts = items.map(d => ({
        ...d,
        x: typeMap[d.type] !== undefined ? typeMap[d.type] : 2.5, // Default for unknown types
        y: d.maxScore,
        z: (d.scoringRate / 100) * RATE_SCALE
    }));
    
    // Default: All categories active
    if (vizState.activeCats.size === 0) {
        globalData.categories.forEach(c => vizState.activeCats.add(c));
    }
    
    // Initialize Category Buttons and Legend
    renderVizFilters();
    
    resizeViz();
    drawViz();
}

function renderVizFilters() {
    const btnRow = document.getElementById("viz-btn-row");
    const legend = document.getElementById("viz-legend");
    btnRow.innerHTML = '';
    legend.innerHTML = '';
    
    const allCats = globalData.categories;
    const catsForBtn = ["전체", ...allCats];
    
    catsForBtn.forEach(cat => {
        // Color mapping for new categories
        if (cat !== "전체" && !vizState.catColors[cat]) {
            const hue = Math.floor(Math.random() * 360);
            vizState.catColors[cat] = `hsl(${hue}, 60%, 50%)`;
        }
        
        const btn = document.createElement("button");
        btn.className = "viz-cat-btn";
        
        // Active status logic
        updateBtnStatus(btn, cat);

        if (cat !== "전체") {
            const dot = document.createElement("span");
            dot.style.cssText = `display:inline-block;width:8px;height:8px;border-radius:50%;background:${vizState.catColors[cat]};margin-right:5px;vertical-align:middle;`;
            btn.appendChild(dot);
        }
        btn.appendChild(document.createTextNode(cat));
        
        btn.addEventListener("click", () => {
            if (cat === "전체") {
                const isAllActive = allCats.every(c => vizState.activeCats.has(c));
                if (isAllActive) {
                    vizState.activeCats.clear();
                } else {
                    allCats.forEach(c => vizState.activeCats.add(c));
                }
            } else {
                // Toggle specific category
                if (vizState.activeCats.has(cat)) {
                    vizState.activeCats.delete(cat);
                } else {
                    vizState.activeCats.add(cat);
                }
            }
            // Refresh all buttons UI
            document.querySelectorAll(".viz-cat-btn").forEach(b => {
                const bText = b.textContent.trim();
                updateBtnStatus(b, bText);
            });
            drawViz();
        });
        btnRow.appendChild(btn);
        
        if (cat !== "전체") {
            const legItem = document.createElement("div");
            legItem.className = "leg-item";
            legItem.innerHTML = `<span class="leg-dot" style="background:${vizState.catColors[cat]}"></span><span>${cat}</span>`;
            legend.appendChild(legItem);
        }
    });
}

function updateBtnStatus(btn, catName) {
    const allCats = globalData.categories;
    if (catName === "전체") {
        // "전체" button is active if ALL categories are active
        const isAllActive = allCats.every(c => vizState.activeCats.has(c));
        if (isAllActive) btn.classList.add("active");
        else btn.classList.remove("active");
    } else {
        if (vizState.activeCats.has(catName)) btn.classList.add("active");
        else btn.classList.remove("active");
    }
}

function resizeViz() {
    const dpr = window.devicePixelRatio || 1;
    const rect = vizState.canvas.getBoundingClientRect();
    vizState.width = rect.width;
    vizState.height = 550; // Significantly reduced height for compact view
    vizState.canvas.width = vizState.width * dpr;
    vizState.canvas.height = vizState.height * dpr;
    vizState.ctx.scale(dpr, dpr);
}

function project(x3, y3, z3) {
    const xmin=-1, xmax=2, ymin=0, ymax=7, zmin=0, zmax=vizState.RATE_SCALE;
    
    const norm = (v, mn, mx) => (v - mn) / (mx - mn);
    const nx = (norm(x3, xmin, xmax) - 0.5) * 2;
    const ny = (norm(y3, ymin, ymax) - 0.5) * 2;
    const nz = (norm(z3, zmin, zmax) - 0.5) * 2;

    const cosY = Math.cos(vizState.rotY), sinY = Math.sin(vizState.rotY);
    const rx1 = nx * cosY - nz * sinY; 
    const ry1 = ny;
    const rz1 = nx * sinY + nz * cosY;

    const cosX = Math.cos(vizState.rotX), sinX = Math.sin(vizState.rotX);
    const rx2 = rx1;
    const ry2 = ry1 * cosX - rz1 * sinX; 
    const rz2 = ry1 * sinX + rz1 * cosX;

    const fov = 10.0;
    const scale = fov / (fov + rz2);
    
    // vizState.zoom을 곱하여 실제 휠 조작이 반영되도록 수정
    const cubeScale = Math.min(vizState.width * 0.15, vizState.height * 0.3) * vizState.zoom;
    const px = vizState.width / 2 + rx2 * scale * cubeScale * 1.5;
    const py = vizState.height * 0.55 - ry2 * scale * cubeScale * 1.25; 

    return { px, py, depth: rz2 };
}

function drawViz() {
    const ctx = vizState.ctx;
    const CW = vizState.width;
    const CH = vizState.height;
    
    ctx.clearRect(0, 0, CW, CH);
    
    // Draw Axis
    const axisColor = "rgba(100,100,100,0.2)";
    const lblColor = "#64748b";
    
    ctx.save();
    ctx.strokeStyle = axisColor;
    ctx.lineWidth = 0.8;
    ctx.setLineDash([4, 4]);

    const xmin=-1, xmax=2, ymin=0, ymax=7, zmin=0, zmax=vizState.RATE_SCALE;

    // X Grids
    [-1, 0, 1, 2].forEach(xv => {
        const p0 = project(xv, ymin, zmin);
        const p1 = project(xv, ymax, zmin);
        const p2 = project(xv, ymin, zmax);
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p1.px, p1.py); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p2.px, p2.py); ctx.stroke();
    });

    // Y Grids
    [0, 2, 3, 5, 7].forEach(yv => {
        const p0 = project(xmin, yv, zmin);
        const p1 = project(xmax, yv, zmin);
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p1.px, p1.py); ctx.stroke();
    });

    // Z Grids (X-Z plane base)
    [0, 0.2, 0.4, 0.6, 0.8, 1.0].forEach(rv => {
        const p0 = project(xmin, ymin, rv * vizState.RATE_SCALE);
        const p1 = project(xmax, ymin, rv * vizState.RATE_SCALE);
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p1.px, p1.py); ctx.stroke();
    });

    // Y-Z Plane Grids (Back wall)
    [0, 2, 3, 5, 7].forEach(yv => {
        const p0 = project(xmin, yv, zmin);
        const p1 = project(xmin, yv, zmax);
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p1.px, p1.py); ctx.stroke();
    });
    [0, 0.2, 0.4, 0.6, 0.8, 1.0].forEach(rv => {
        const p0 = project(xmin, ymin, rv * vizState.RATE_SCALE);
        const p1 = project(xmin, ymax, rv * vizState.RATE_SCALE);
        ctx.beginPath(); ctx.moveTo(p0.px, p0.py); ctx.lineTo(p1.px, p1.py); ctx.stroke();
    });

    // Main Axes
    ctx.setLineDash([]);
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = "rgba(100,100,100,0.4)";

    const O = project(xmin, ymin, zmin);
    const Xend = project(xmax + 0.3, ymin, zmin);
    const Yend = project(xmin, ymax + 0.8, zmin);
    const Zend = project(xmin, ymin, zmax + 0.8);

    [[O, Xend], [O, Yend], [O, Zend]].forEach(([a, b]) => {
        ctx.beginPath(); ctx.moveTo(a.px, a.py); ctx.lineTo(b.px, b.py); ctx.stroke();
    });

    // Axis Labels
    ctx.fillStyle = lblColor;
    ctx.font = "11px Pretendard, sans-serif";
    ctx.textAlign = "center";
    
    const xl = project(xmax + 0.6, ymin, zmin);
    ctx.fillText("유형(X)", xl.px, xl.py + 4);
    const yl = project(xmin, ymax + 0.9, zmin);
    ctx.fillText("만점(Y)", yl.px - 15, yl.py - 3);
    const zl = project(xmin, ymin, zmax + 0.9);
    ctx.fillText("득점률(Z)", zl.px - 20, zl.py - 3);

    // X Ticks
    const xLabels = {"-1":"정책", "0":"목표", "1":"위험관리", "2":"성과"};
    [-1, 0, 1, 2].forEach(xv => {
        const p = project(xv, ymin, zmin);
        ctx.fillText(xLabels[String(xv)] || "", p.px, p.py + 14);
    });

    // Y Ticks
    [2, 3, 5, 7].forEach(yv => {
        const p = project(xmin - 0.1, yv, zmin);
        ctx.textAlign = "right";
        ctx.fillText(yv + "점", p.px - 4, p.py + 3);
    });

    // Z Ticks
    [0, 0.2, 0.4, 0.6, 0.8, 1.0].forEach(rv => {
        const p = project(xmin, ymin, rv * vizState.RATE_SCALE);
        ctx.textAlign = "right";
        ctx.fillText((rv * 100).toFixed(0) + "%", p.px - 4, p.py + 3);
    });
    
    ctx.restore();

    // Sort by depth
    const projected = vizState.pts.map(d => {
        const { px, py, depth } = project(d.x, d.y, d.z);
        return { ...d, px, py, depth };
    }).sort((a, b) => b.depth - a.depth);

    // Draw active points (Hide inactive ones completely)
    projected.filter(d => vizState.activeCats.has(d.category)).forEach(d => {
        const isHov = vizState.hoveredPt && vizState.hoveredPt.id === d.id;
        const r = isHov ? 8 : 5;
        ctx.save();
        ctx.globalAlpha = isHov ? 1 : 0.8;
        ctx.beginPath();
        ctx.arc(d.px, d.py, r, 0, Math.PI * 2);
        ctx.fillStyle = vizState.catColors[d.category] || "#888";
        ctx.fill();
        if (isHov) {
            ctx.strokeStyle = "#fff";
            ctx.lineWidth = 2;
            ctx.stroke();
            
            ctx.fillStyle = "#000";
            ctx.font = "bold 11px Pretendard, sans-serif";
            ctx.textAlign = "center";
            ctx.fillText("#" + d.id, d.px, d.py - 12);
        }
        ctx.restore();
    });
}

function handleVizMouseMove(e) {
    if (currentViewMode !== 'VIZ') return;
    
    if (vizState.dragOn) {
        vizState.rotY += (e.clientX - vizState.lastMx) * 0.008;
        vizState.rotX += (e.clientY - vizState.lastMy) * 0.008;
        vizState.lastMx = e.clientX; 
        vizState.lastMy = e.clientY;
        drawViz();
    }

    const rect = vizState.canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    const vis = vizState.pts.filter(d => vizState.activeCats.has(d.category));

    let found = null, minD = Infinity;
    vis.forEach(d => {
        const { px, py } = project(d.x, d.y, d.z);
        const dist = Math.hypot(px - mx, py - my);
        if (dist < 15 && dist < minD) { minD = dist; found = { ...d, px, py }; }
    });

    if (found !== vizState.hoveredPt) { 
        vizState.hoveredPt = found; 
        drawViz(); 
    }

    const tooltip = document.getElementById("viz-tooltip");
    if (found) {
        tooltip.style.display = "block";
        tooltip.style.left = (found.px + 12) + "px";
        tooltip.style.top = Math.max(0, found.py - 80) + "px";
        tooltip.innerHTML =
            `<b>#${found.id}</b><br>` +
            `${found.category}<br>` +
            `유형: ${found.type}<br>` +
            `만점: ${found.maxScore}점<br>` +
            `득점률: ${found.scoringRate.toFixed(1)}%`;
    } else {
        tooltip.style.display = "none";
    }
}

// Criteria Modal Handlers
window.openCriteriaModal = function() {
    const modal = document.getElementById('modal-criteria');
    if (modal) modal.style.display = 'flex';
}

window.closeCriteriaModal = function() {
    const modal = document.getElementById('modal-criteria');
    if (modal) modal.style.display = 'none';
}
