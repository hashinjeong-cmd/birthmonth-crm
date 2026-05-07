// 주요 전역 상태
let birthdays = [];

document.addEventListener('DOMContentLoaded', () => {
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
            saveToLocalStorage();
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

    loadFromLocalStorage();
    renderAll();
});

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
        saveToLocalStorage();
        renderAll();
    }
};

window.clearMonth = function(month) {
    if (confirm(`${month}월 대상자 목록을 모두 비우시겠습니까?`)) {
        birthdays = birthdays.filter(b => b.month !== month);
        saveToLocalStorage();
        renderAll();
    }
};

/* 엑셀 파싱 로직 (오류 수정 및 안정성 강화) */
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            // 파일을 더 안정적으로 읽어오도록 수정
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
                birthdays = [...birthdays, ...parsedData];
                saveToLocalStorage();
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
        // [수정 완료] 아래 줄에서 '소속', '부서'에 따옴표가 빠져서 에러가 나던 것을 고쳤습니다!
        let branchVal = row['지점명'] || row['지점'] || row['소속'] || row['부서'];
        let noteVal = row['
