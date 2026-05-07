// 주요 전역 상태
let birthdays = [];

document.addEventListener('DOMContentLoaded', () => {
    // 1. 페이지가 열리자마자 저장된 데이터를 불러옵니다.
    loadFromLocalStorage();
    
    const dlBtns = ['hero-sample-download', 'dash-sample-download'];
    dlBtns.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) btn.addEventListener('click', downloadSampleExcel);
    });

    const uploadInputs = ['excel-upload-main', 'excel-upload-dash'];
    uploadInputs.forEach(id => {
        const input = document.getElementById(id);
        if (input) input.addEventListener('change', handleExcelUpload);
    });

    document.getElementById('clear-data').addEventListener('click', () => {
        if (confirm("정말 모든 데이터를 삭제하시겠습니까? (되돌릴 수 없습니다)")) {
            birthdays = [];
            saveToLocalStorage(); // 빈 배열 저장 (전체 삭제)
            renderAll();
            alert("전체 데이터가 삭제되었습니다.");
        }
    });

    document.getElementById('lunar-modal').addEventListener('click', (e) => {
        if (e.target === document.getElementById('lunar-modal')) closeLunarModal();
    });

    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            renderDashboard(e.target.value);
        });
    }

    // 2. 불러온 데이터를 화면에 그립니다.
    renderAll();
});

/* 로컬 스토리지 저장 및 로드 함수 */
function saveToLocalStorage() {
    // 'birthTrakData'라는 이름으로 브라우저에 데이터를 영구 저장합니다.
    localStorage.setItem('birthTrakData', JSON.stringify(birthdays));
}

function loadFromLocalStorage() {
    try {
        const saved = localStorage.getItem('birthTrakData');
        birthdays = saved ? JSON.parse(saved) : [];
    } catch (e) {
        console.error("데이터 로딩 실패:", e);
        birthdays = [];
    }
}

/* Modal (음력 변환) 함수들 */
function openLunarModal() {
    document.getElementById('lunar-m').value = '';
    document.getElementById('lunar-d').value = '';
    document.getElementById('lunar-result').style.display = 'none';
    document.getElementById('lunar-modal').style.display = 'flex';
}

function closeLunarModal() {
    document.getElementById('lunar-modal').style.display = 'none';
}

function convertLunarToSolar() {
    const m = parseInt(document.getElementById('lunar-m').value);
    const d = parseInt(document.getElementById('lunar-d').value);
    const resultBox = document.getElementById('lunar-result');
    
    if (!m || !d || m < 1 || m > 12 || d < 1 || d > 31) {
        alert("유효한 음력 월, 일을 입력해주세요.");
        return;
    }

    try {
        const year = new Date().getFullYear();
        const lunarDate = Lunar.fromYmd(year, m, d);
        const solarDate = lunarDate.getSolar();
        
        const sm = solarDate.getMonth();
        const sd = solarDate.getDay();

        resultBox.innerHTML = `
            <p>올해(${year}년)의 양력 날짜는</p>
            <div class="highlight">${sm}월 ${sd}일</div>
        `;
        resultBox.style.display = 'block';
    } catch (e) {
        console.error(e);
        resultBox.innerHTML = `<p style="color:var(--danger)">계산 중 오류가 발생했습니다.</p>`;
        resultBox.style.display = 'block';
    }
}

/* 카카오톡 축하 공유 함수 */
function shareToKakao(name, branch) {
    const branchText = branch ? `${branch} ` : '';
    const defaultText = `🎉 ${branchText}${name}님의 생일을 진심으로 축하드립니다!\n오늘 하루 행복한 일 가득하시고, 늘 건강과 좋은 일만 함께하시길 바랍니다 😊`;
    
    if (navigator.share) {
        navigator.share({
            title: '생일 축하 메시지',
            text: defaultText,
        }).catch(err => {
            console.log("공유 취소됨: ", err);
        });
    } else {
        navigator.clipboard.writeText(defaultText).then(() => {
            alert(`"${name}" 님을 위한 축하 메시지가 복사되었습니다!\n카카오톡 대화창에 붙여넣기(Ctrl+V) 해주세요.`);
        }).catch(() => {
            alert("공유를 지원하지 않는 브라우저입니다.");
        });
    }
}

/* 개별 대상자 및 월별 대상자 삭제 기능 */
window.deletePerson = function(id) {
    if (confirm("해당 대상자를 삭제하시겠습니까?")) {
        birthdays = birthdays.filter(b => b.id !== id);
        saveToLocalStorage(); // 변경사항 저장
        renderAll();
    }
};

window.clearMonth = function(month) {
    if (confirm(`${month}월 대상자 목록을 모두 비우시겠습니까?`)) {
        birthdays = birthdays.filter(b => b.month !== month);
        saveToLocalStorage(); // 변경사항 저장
        renderAll();
    }
};

/* 엑셀 파싱 로직 */
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            
            if (json.length === 0) {
                alert("엑셀 파일이 비어있습니다.");
                return;
            }

            const parsedData = processExcelData(json);
            if (parsedData.length > 0) {
                // 기존 데이터에 새 데이터를 합칩니다.
                birthdays = [...birthdays, ...parsedData];
                saveToLocalStorage(); // 합친 데이터를 저장합니다.
                renderAll();
                alert(`${parsedData.length}건의 데이터 연동 성공!`);
                document.getElementById('dashboard').scrollIntoView({behavior: 'smooth'});
            } else {
                alert("데이터를 찾을 수 없습니다. 양식의 이름(첫 줄)이 맞는지 확인해 주세요.");
            }
        } catch (error) {
            console.error(error);
            alert("파일을 읽는 중 오류가 발생했습니다.");
        }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
}

function processExcelData(json) {
    const results = [];
    json.forEach(row => {
        let name = row['이름'] || row['성명'] || row['Name'] || row['name'];
        let dateVal = row['생일'] || row['생년월일'] || row['양력생일'] || row['날짜'];
        let branchVal = row['지점명'] || row['지점'] || row['소속'] || row['부서'];
        let noteVal = row['특이사항'] || row['비고'] || row['메모'];
        
        if (name && dateVal) {
            const parsedDate = parseDateString(dateVal);
            if (parsedDate) {
                results.push({
                    id: Math.random().toString(36).substring(7),
                    name: String(name).trim(),
                    month: parsedDate.month,
                    day: parsedDate.day,
                    branch: branchVal ? String(branchVal).trim() : '',
                    note: noteVal ? String(noteVal).trim() : ''
                });
            }
        }
    });
    return results;
}

function parseDateString(val) {
    let m = null, d = null;
    if (typeof val === 'number') {
        const excelEpoch = new Date(1899, 11, 31);
        const dateObj = new Date(excelEpoch.getTime() + val * 86400000);
        m = dateObj.getMonth() + 1;
        d = dateObj.getDate();
    } else if (typeof val === 'string') {
        const matchRegex = /(\d{1,4})[./-](\d{1,2})[./-](\d{1,2})|(\d{1,2})[./-](\d{1,2})/;
        const match = val.match(matchRegex);
        if (match) {
            if (match[1]) { m = parseInt(match[2], 10); d = parseInt(match[3], 10); } 
            else { m = parseInt(match[4], 10); d = parseInt(match[5], 10); }
        } else {
            const jsDate = new Date(val);
            if (!isNaN(jsDate.getTime())) { m = jsDate.getMonth() + 1; d = jsDate.getDate(); }
        }
    }
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) return { month: m, day: d };
    return null;
}

function renderAll() {
    renderCurrentMonthHero();
    renderDashboard();
}

function renderCurrentMonthHero() {
    const listContainer = document.getElementById('hero-birthday-list');
    const badge = document.getElementById('current-month-count');
    const alertCard = document.getElementById('hero-alert-card');
    const today = new Date();
    const currentMonth = today.getMonth() + 1;
    const currentDay = today.getDate();

    if (birthdays.length === 0) {
        listContainer.innerHTML = `<div class="empty-state-mini">데이터를 업로드하면 당월 생일자가 표시됩니다.</div>`;
        badge.innerText = '대기중';
        alertCard.style.display = 'none';
        return;
    }

    let thisMonthGroups = birthdays.filter(b => b.month === currentMonth).sort((a,b) => a.day - b.day);

    if (thisMonthGroups.length === 0) {
        listContainer.innerHTML = `<div class="empty-state-mini">이번 달 생일자가 없습니다.</div>`;
        badge.innerText = '이번 달 0명';
        alertCard.style.display = 'none';
        return;
    }

    badge.innerText = `이번 달 ${thisMonthGroups.length}명`;
    listContainer.innerHTML = '';
    
    let closestDays = 999;
    let closestUpcoming = null;

    thisMonthGroups.forEach((b, index) => {
        let dDayText = '';
        const diff = b.day - currentDay;
        
        if (diff === 0) dDayText = '오늘! 🎉';
        else if (diff === 1) dDayText = '내일';
        else if (diff > 1) dDayText = `D-${diff}`;
        else dDayText = `지남 (${b.month}/${b.day})`;

        if (diff >= 0 && diff < closestDays) {
            closestDays = diff;
            closestUpcoming = b;
        }
        
        let branchTag = b.branch ? `<span class="branch-badge">${b.branch}</span>` : '';
        let noteText = b.note ? `<span class="note-text"> ${b.note}</span>` : '';

        const li = document.createElement('li');
        li.className = 'birthday-item';
        li.innerHTML = `
            <div class="avatar gradient-${index % 5}">${b.name.charAt(0)}</div>
            <div class="info">
                <h4>${b.name}${branchTag}</h4>
                <p>${dDayText}${noteText}</p>
            </div>
            <div class="action flex" style="display:flex; gap:0.5rem; align-items:center;">
                <button class="btn-icon-kakao share-btn" data-name="${b.name}" data-branch="${b.branch}" title="카톡 메시지 공유">💬</button>
                <button class="btn-icon-danger" onclick="deletePerson('${b.id}')" title="대상자 삭제">✕</button>
            </div>
        `;
        listContainer.appendChild(li);
    });

    if (closestUpcoming && closestDays <= 7) {
        alertCard.style.display = 'flex';
        let alertBranch = closestUpcoming.branch ? `[${closestUpcoming.branch}] ` : '';
        document.getElementById('hero-alert-text').innerText = `${alertBranch}${closestUpcoming.name}님 생일이 ${closestDays === 0 ? '오늘입니다!' : '다가옵니다!'}`;
    } else {
        alertCard.style.display = 'none';
    }

    attachShareEvents(listContainer);
}

function renderDashboard(query = '') {
    const dashContainer = document.getElementById('dashboard-content');
    
    let currentData = birthdays;
    query = query.trim().toLowerCase();
    
    if (query) {
        currentData = birthdays.filter(b => 
            b.name.toLowerCase().includes(query) || 
            (b.branch && b.branch.toLowerCase().includes(query))
        );
    }

    if (birthdays.length === 0) {
        dashContainer.innerHTML = `
            <div class="empty-state">
                <h3>아직 등록된 데이터가 없습니다.</h3>
                <p>상단의 '데이터 업로드' 버튼을 통해 엑셀 파일을 업로드 해주세요.</p>
                <p class="guide-text">※ 처음이신가요? <strong>'양식 다운로드'</strong> 버튼을 통해 템플릿을 받아 작성해 보세요.</p>
            </div>`;
        return;
    }
    
    if (query && currentData.length === 0) {
        dashContainer.innerHTML = `
            <div class="empty-state">
                <h3>검색 결과가 없습니다 😢</h3>
                <p>"${query}"에 해당하는 대상자를 찾을 수 없습니다.</p>
            </div>`;
        return;
    }

    let gridHtml = '<div class="months-grid">';
    const currentMonth = new Date().getMonth() + 1;

    for (let m = 1; m <= 12; m++) {
        const mBirthdays = currentData.filter(b => b.month === m).sort((a,b) => a.day - b.day);
        
        if (query && mBirthdays.length === 0) continue;

        let highlightStyle = (m === currentMonth && !query) ? 'border: 1px solid var(--accent-1); box-shadow: 0 0 20px rgba(236, 72, 153, 0.2);' : '';
        
        gridHtml += `
            <div class="month-card" style="${highlightStyle}">
                <div class="month-title">
                    <div>
                        <span>${m}월 ${m === currentMonth ? '✨' : ''}</span>
                        <span class="count">${mBirthdays.length}명</span>
                    </div>
                    ${mBirthdays.length > 0 ? `<button class="btn-clear-month" onclick="clearMonth(${m})" title="${m}월 대상자 모두 삭제">비우기</button>` : ''}
                </div>
                <ul class="birthday-list" style="max-height: 250px;">
        `;
        if (mBirthdays.length === 0) {
            gridHtml += `<li style="color: var(--text-secondary); font-size: 0.9rem; text-align: center; padding: 1rem 0;">데이터 없음</li>`;
        } else {
            mBirthdays.forEach((b, idx) => {
                let branchTag = b.branch ? `<span class="branch-badge mini">${b.branch}</span>` : '';
                let noteText = b.note ? `<span class="note-text"> ${b.note}</span>` : '';

                gridHtml += `
                    <li class="birthday-item" style="padding: 0.7rem; margin-bottom: 0.5rem;">
                        <div class="avatar gradient-${idx % 5}" style="width:36px; height:36px; font-size:1rem; border-radius:10px;">${b.name.charAt(0)}</div>
                        <div class="info">
                            <h4 style="font-size: 0.95rem; margin-bottom: 0.2rem;">${b.name}${branchTag}</h4>
                            <p style="font-size: 0.8rem; color: var(--text-secondary);">${m}월 ${b.day}일${noteText}</p>
                        </div>
                        <div style="display:flex; gap:0.5rem; justify-content:center; align-items:center;">
                            <button class="btn-icon-kakao share-btn" style="width:30px; height:30px; font-size:0.85rem;" data-name="${b.name}" data-branch="${b.branch}">💬</button>
                            <button class="btn-icon-danger" style="width:30px; height:30px; font-size:0.85rem;" onclick="deletePerson('${b.id}')">✕</button>
                        </div>
                    </li>
                `;
            });
        }
        gridHtml += `</ul></div>`;
    }
    gridHtml += '</div>';
    dashContainer.innerHTML = gridHtml;
    
    attachShareEvents(dashContainer);
}

function attachShareEvents(container) {
    const shareBtns = container.querySelectorAll('.share-btn');
    shareBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const name = e.currentTarget.getAttribute('data-name');
            const branch = e.currentTarget.getAttribute('data-branch');
            shareToKakao(name, branch);
        });
    });
}

function downloadSampleExcel() {
    const ws_data = [
        ["이름", "생일", "지점명", "특이사항"],
        ["홍길동", "1999-01-01", "서울지점", "AMC"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    
    const wscols = [ {wch: 10}, {wch: 15}, {wch: 15}, {wch: 35} ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "지점별생일목록");
    XLSX.writeFile(wb, "지점별_생일관리_양식.xlsx");
}
