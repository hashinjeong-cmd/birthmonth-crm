// 주요 전역 상태
let birthdays = [];

document.addEventListener('DOMContentLoaded', () => {
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

    const clearBtn = document.getElementById('clear-data');
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            if (confirm("정말 모든 데이터를 삭제하시겠습니까? (되돌릴 수 없습니다)")) {
                birthdays = [];
                saveToLocalStorage();
                renderAll();
                alert("전체 데이터가 삭제되었습니다.");
            }
        });
    }

    const lunarModal = document.getElementById('lunar-modal');
    if (lunarModal) {
        lunarModal.addEventListener('click', (e) => {
            if (e.target === lunarModal) closeLunarModal();
        });
    }

    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            renderDashboard(e.target.value);
        });
    }

    renderAll();
});

function saveToLocalStorage() {
    try {
        localStorage.setItem('birthTrakData', JSON.stringify(birthdays));
    } catch (e) {
        console.error("저장 실패", e);
    }
}

function loadFromLocalStorage() {
    try {
        const saved = localStorage.getItem('birthTrakData');
        birthdays = saved ? JSON.parse(saved) : [];
    } catch (e) {
        birthdays = [];
    }
}

/* Modal 함수 */
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
        
        resultBox.innerHTML = `
            <p>올해(${year}년)의 양력 날짜는</p>
            <div class="highlight">${solarDate.getMonth()}월 ${solarDate.getDay()}일</div>
        `;
        resultBox.style.display = 'block';
    } catch (e) {
        resultBox.innerHTML = `<p style="color:var(--danger)">변환 중 오류가 발생했습니다.</p>`;
        resultBox.style.display = 'block';
    }
}

/* 카카오톡 공유 */
function shareToKakao(name, branch) {
    const branchText = branch ? `${branch} ` : '';
    const defaultText = `🎉 ${branchText}${name}님의 생일을 진심으로 축하드립니다!\n오늘 하루 행복한 일 가득하시고, 늘 건강과 좋은 일만 함께하시길 바랍니다 😊`;
    
    if (navigator.share) {
        navigator.share({ title: '생일 축하', text: defaultText }).catch(() => {});
    } else {
        navigator.clipboard.writeText(defaultText).then(() => {
            alert("메시지가 복사되었습니다! 카톡에 붙여넣어 주세요.");
        });
    }
}

window.deletePerson = function(id) {
    if (confirm("삭제하시겠습니까?")) {
        birthdays = birthdays.filter(b => b.id !== id);
        saveToLocalStorage();
        renderAll();
    }
};

window.clearMonth = function(month) {
    if (confirm(`${month}월 데이터를 모두 비울까요?`)) {
        birthdays = birthdays.filter(b => b.month !== month);
        saveToLocalStorage();
        renderAll();
    }
};

/* 엑셀 업로드 (모바일 최적화) */
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            // 모바일 호환성을 위해 ArrayBuffer 방식으로 읽기
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(worksheet);
            
            if (json.length === 0) {
                alert("파일이 비어있습니다.");
                return;
            }

            const parsedData = processExcelData(json);
            if (parsedData.length > 0) {
                birthdays = [...birthdays, ...parsedData];
                saveToLocalStorage();
                renderAll();
                alert(`${parsedData.length}건 연동 성공!`);
            } else {
                alert("데이터를 찾을 수 없습니다. 양식을 확인해 주세요.");
            }
        } catch (error) {
            alert("파일 분석 중 오류가 발생했습니다.");
        }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
}

function processExcelData(json) {
    const results = [];
    const currentYear = new Date().getFullYear();

    json.forEach(row => {
        let name = row['이름'] || row['성명'] || row['Name'] || row['name'];
        let dateVal = row['생일'] || row['생년월일'] || row['양력생일'] || row['음력생일'] || row['날짜'];
        let branchVal = row['지점명'] || row['지점'] || row['소속'] || row['부서']; // 따옴표 추가 완료
        let noteVal = row['특이사항'] || row['비고'] || row['메모'];
        let typeVal = row['구분'] || row['음양'];

        let isLunar = false;
        if (typeVal && String(typeVal).includes('음')) isLunar = true;
        if (noteVal && String(noteVal).includes('음력')) isLunar = true;
        if (dateVal && String(dateVal).includes('음')) isLunar = true;

        if (name && dateVal) {
            const parsedDate = parseDateString(dateVal);
            if (parsedDate) {
                let m = parsedDate.month;
                let d = parsedDate.day;
                let note = noteVal ? String(noteVal).trim() : '';

                if (isLunar) {
                    try {
                        const lunar = Lunar.fromYmd(currentYear, m, d);
                        const solar = lunar.getSolar();
                        const lunarMark = `(음력 ${m}/${d})`;
                        m = solar.getMonth();
                        d = solar.getDay();
                        note = note ? `${note} ${lunarMark}` : lunarMark;
                    } catch (err) {}
                }

                results.push({
                    id: Math.random().toString(36).substring(7),
                    name: String(name).trim(),
                    month: m,
                    day: d,
                    branch: branchVal ? String(branchVal).trim() : '',
                    note: note
                });
            }
        }
    });
    return results;
}

function parseDateString(val) {
    let m = null, d = null;
    if (typeof val === 'number') {
        const dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
        m = dateObj.getUTCMonth() + 1;
        d = dateObj.getUTCDate();
    } else if (typeof val === 'string') {
        const str = val.replace(/\s+/g, '');
        const koMatch = str.match(/(\d{1,2})월(\d{1,2})일/);
        if (koMatch) {
            m = parseInt(koMatch[1]); d = parseInt(koMatch[2]);
        } else {
            const match = str.match(/(\d{1,4})[./-](\d{1,2})[./-](\d{1,2})|(\d{1,2})[./-](\d{1,2})/);
            if (match) {
                if (match[1]) { m = parseInt(match[2]); d = parseInt(match[3]); }
                else { m = parseInt(match[4]); d = parseInt(match[5]); }
            }
        }
    }
    return (m >= 1 && m <= 12 && d >= 1 && d <= 31) ? { month: m, day: d } : null;
}

function renderAll() {
    renderCurrentMonthHero();
    renderDashboard();
}

function renderCurrentMonthHero() {
    const listContainer = document.getElementById('hero-birthday-list');
    const badge = document.getElementById('current-month-count');
    const today = new Date();
    const curM = today.getMonth() + 1;
    const curD = today.getDate();

    if (birthdays.length === 0) {
        listContainer.innerHTML = `<div class="empty-state-mini">데이터를 업로드해 주세요.</div>`;
        return;
    }

    const items = birthdays.filter(b => b.month === curM).sort((a,b) => a.day - b.day);
    badge.innerText = `이번 달 ${items.length}명`;
    listContainer.innerHTML = '';

    items.forEach((b, i) => {
        const diff = b.day - curD;
        const dDay = diff === 0 ? '오늘! 🎉' : (diff > 0 ? `D-${diff}` : `지남`);
        const branch = b.branch ? `<span class="branch-badge">${b.branch}</span>` : '';
        const note = b.note ? `<span class="note-text">${b.note}</span>` : '';

        const li = document.createElement('li');
        li.className = 'birthday-item';
        li.innerHTML = `
            <div class="avatar gradient-${i % 5}">${b.name.charAt(0)}</div>
            <div class="info">
                <h4>${b.name}${branch}</h4>
                <p>${dDay}${note}</p>
            </div>
            <div class="action" style="display:flex; gap:10px;">
                <button class="btn-icon-kakao share-btn" data-name="${b.name}" data-branch="${b.branch}">💬</button>
                <button class="btn-icon-danger" onclick="deletePerson('${b.id}')">✕</button>
            </div>
        `;
        listContainer.appendChild(li);
    });
    attachShareEvents(listContainer);
}

function renderDashboard(query = '') {
    const container = document.getElementById('dashboard-content');
    if (birthdays.length === 0) return;

    let html = '<div class="months-grid">';
    const curM = new Date().getMonth() + 1;

    for (let m = 1; m <= 12; m++) {
        const mItems = birthdays.filter(b => b.month === m).sort((a,b) => a.day - b.day);
        if (query && !mItems.some(b => b.name.includes(query))) continue;

        html += `
            <div class="month-card" style="${m === curM ? 'border:1px solid var(--accent-1)' : ''}">
                <div class="month-title">
                    <span>${m}월 <small>${mItems.length}명</small></span>
                    ${mItems.length > 0 ? `<button class="btn-clear-month" onclick="clearMonth(${m})">비우기</button>` : ''}
                </div>
                <ul class="birthday-list">
                    ${mItems.map((b, idx) => `
                        <li class="birthday-item" style="padding:0.5rem; margin-bottom:5px;">
                            <div class="avatar gradient-${idx % 5}" style="width:30px; height:30px; font-size:0.8rem;">${b.name.charAt(0)}</div>
                            <div class="info">
                                <h4 style="font-size:0.9rem;">${b.name} <small style="color:var(--text-secondary)">${b.branch}</small></h4>
                                <p style="font-size:0.75rem;">${m}/${b.day} <span class="note-text" style="display:inline; margin:0; padding-left:5px;">${b.note}</span></p>
                            </div>
                            <button class="btn-icon-kakao share-btn" style="width:28px; height:28px; font-size:0.7rem;" data-name="${b.name}" data-branch="${b.branch}">💬</button>
                        </li>
                    `).join('')}
                </ul>
            </div>
        `;
    }
    container.innerHTML = html + '</div>';
    attachShareEvents(container);
}

function attachShareEvents(parent) {
    parent.querySelectorAll('.share-btn').forEach(btn => {
        btn.onclick = (e) => shareToKakao(e.currentTarget.dataset.name, e.currentTarget.dataset.branch);
    });
}

function downloadSampleExcel() {
    const data = [["이름", "생일", "구분", "지점명", "특이사항"], ["홍길동", "05-15", "음력", "스타지점", "축하 메시지 예시"]];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "생일목록");
    XLSX.writeFile(wb, "생일관리_양식.xlsx");
}
