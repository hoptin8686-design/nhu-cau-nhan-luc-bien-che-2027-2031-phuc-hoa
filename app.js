/**
 * HỆ THỐNG XÁC ĐỊNH NHU CẦU NHÂN LỰC VÀ PHƯƠNG ÁN BIÊN CHẾ 2027-2031
 * TRƯỜNG THPT PHỤC HÒA - TỈNH CAO BẰNG
 * JAVASCRIPT APPLICATION LOGIC
 */

document.addEventListener('DOMContentLoaded', () => {
  initCountdown();
  initMainTabs();
  initSubTabs();
  initSearchAndFilter();
  initSimulator();
  initLegalAccordion();
  initCharts();
});

// 1. COUNTDOWN TO SUBMISSION DEADLINE (30/10/2026 17:00)
function initCountdown() {
  const deadline = new Date(2026, 9, 30, 17, 0, 0).getTime(); // Oct 30, 2026
  const countdownEl = document.getElementById('countdown-timer');
  if (!countdownEl) return;

  function update() {
    const now = new Date().getTime();
    const diff = deadline - now;

    if (diff <= 0) {
      countdownEl.innerHTML = '<span style="color:#22c55e; font-weight:700;">HỒ SƠ ĐÃ SẴN SÀNG NỘP SỞ GD&ĐT</span>';
      return;
    }

    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));

    countdownEl.innerHTML = `Hạn nộp Sở GD&ĐT: <strong>${days} ngày ${hours} giờ ${minutes} phút</strong>`;
  }

  update();
  setInterval(update, 60000);
}

// 2. MAIN TABS SWITCHER
function initMainTabs() {
  const tabs = document.querySelectorAll('.nav-tab');
  const panes = document.querySelectorAll('.tab-pane');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetId = tab.getAttribute('data-tab');

      tabs.forEach(t => t.classList.remove('active'));
      panes.forEach(p => p.classList.remove('active'));

      tab.classList.add('active');
      const targetPane = document.getElementById(targetId);
      if (targetPane) {
        targetPane.classList.add('active');
      }

      window.scrollTo({ top: 180, behavior: 'smooth' });
    });
  });
}

// 3. SUB TABS (FOR 9 BIỂU NGHIỆP VỤ)
function initSubTabs() {
  const subTabs = document.querySelectorAll('.sub-tab-btn');
  const subPanes = document.querySelectorAll('.sub-pane');

  subTabs.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-subtab');

      subTabs.forEach(b => b.classList.remove('active'));
      subPanes.forEach(p => p.classList.remove('active'));

      btn.classList.add('active');
      const targetSubPane = document.getElementById(targetId);
      if (targetSubPane) {
        targetSubPane.classList.add('active');
      }
    });
  });
}

// 4. SEARCH & FILTER IN DATA TABLES
function initSearchAndFilter() {
  // Global search input for current active sub-pane table
  const searchInputs = document.querySelectorAll('.table-search-input');

  searchInputs.forEach(input => {
    input.addEventListener('input', (e) => {
      const term = e.target.value.toLowerCase().trim();
      const tableId = input.getAttribute('data-target-table');
      const table = document.getElementById(tableId);
      if (!table) return;

      const rows = table.querySelectorAll('tbody tr');
      rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(term)) {
          row.style.display = '';
        } else {
          row.style.display = 'none';
        }
      });
    });
  });
}

// 5. INTERACTIVE SIMULATOR (TÍNH TOÁN BIÊN CHẾ ĐỊNH MỨC THEO TT 20/2023 VÀ HD 07)
function initSimulator() {
  const classSlider = document.getElementById('sim-class-slider');
  const classValEl = document.getElementById('sim-class-val');
  const regionSelect = document.getElementById('sim-region-select');

  if (!classSlider || !classValEl) return;

  function recalculate() {
    const classes = parseInt(classSlider.value, 10);
    const regionRatio = parseFloat(regionSelect.value); // 2.45 for Vùng 1
    classValEl.textContent = `${classes} lớp`;

    // 1. Quỹ học sinh ước tính (~41 HS/lớp)
    const estStudents = classes * 41;
    document.getElementById('sim-students-res').textContent = `${estStudents} HS`;

    // 2. Giáo viên định mức theo Thông tư 20/2023/TT-BGDĐT
    const gvQuota = (classes * regionRatio).toFixed(1);
    document.getElementById('sim-gv-quota-res').textContent = `${gvQuota} GV`;

    // 3. Số người quy đổi PP2 (FTE) theo quỹ giờ:
    // Bình quân mỗi lớp học THPT có ~29 tiết chuyên môn/tuần * 35 tuần = 1.015 tiết * 45p = 761h dạy + 700h chuẩn bị/đánh giá/trải nghiệm = ~1.460h ròng/lớp
    // Tổng giờ ròng GV = classes * 1.460h. Quỹ giờ chuẩn GV = 1.470h => FTE GV ≈ classes * 0.99
    // Cộng BGH (3 người x 1.904h), Nhân viên (6 người x 1.904h)
    const totalGvHours = classes * 1460;
    const totalAdminHours = 9 * 1904; // 3 BGH + 6 NV
    const totalHours = totalGvHours + totalAdminHours;
    const ftePP2 = ((totalGvHours / 1470) + 9).toFixed(2);
    document.getElementById('sim-fte-res').textContent = `${ftePP2} người`;

    // 4. Số người cần bố trí kiến nghị
    // Đối với trường THPT: 1 HT + 2 PHT (nếu >=10 lớp) + số GV làm tròn + 6 NV + 2 HĐ 111
    const pht = classes >= 10 ? 2 : 1;
    const gvCeil = Math.ceil(parseFloat(gvQuota));
    const totalStaff = 1 + pht + gvCeil + 6 + 2;
    const vcStaff = 1 + pht + gvCeil + 6;

    document.getElementById('sim-total-staff-res').textContent = `${totalStaff} người`;
    document.getElementById('sim-vc-staff-res').textContent = `${vcStaff} biên chế`;

    // Đánh giá so với hiện có (35 người)
    const diff = totalStaff - 35;
    const diffEl = document.getElementById('sim-diff-res');
    if (diff > 0) {
      diffEl.innerHTML = `<span style="color:#059669; font-weight:700;">+${diff} người (Cần đề xuất bổ sung ${diff} biên chế GV)</span>`;
    } else if (diff === 0) {
      diffEl.innerHTML = `<span style="color:#0284c7; font-weight:700;">Bằng đúng biên chế hiện có (35 người)</span>`;
    } else {
      diffEl.innerHTML = `<span style="color:#eab308; font-weight:700;">Dư ${Math.abs(diff)} người so với định mức</span>`;
    }
  }

  classSlider.addEventListener('input', recalculate);
  regionSelect.addEventListener('change', recalculate);
  recalculate();
}

// 6. LEGAL ACCORDION
function initLegalAccordion() {
  const headers = document.querySelectorAll('.accordion-header');
  headers.forEach(h => {
    h.addEventListener('click', () => {
      const item = h.parentElement;
      item.classList.toggle('open');
    });
  });
}

// 7. CHARTS RENDERING (CHART.JS)
function initCharts() {
  // Chart 1: Personnel Distribution
  const ctxStaff = document.getElementById('chart-staff-distribution');
  if (ctxStaff && window.Chart) {
    new Chart(ctxStaff, {
      type: 'doughnut',
      data: {
        labels: [
          'Lãnh đạo Quản lý (3)',
          'Giáo viên Tự nhiên (13)',
          'Giáo viên Xã hội (11)',
          'Nhân viên hỗ trợ (6)',
          'Hợp đồng 111 (2)'
        ],
        datasets: [{
          data: [3, 13, 11, 6, 2],
          backgroundColor: [
            '#1b365d',
            '#0891b2',
            '#059669',
            '#f59e0b',
            '#8b5cf6'
          ],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: 12,
              font: { family: 'Inter', size: 12 }
            }
          }
        },
        cutout: '65%'
      }
    });
  }

  // Chart 2: 5-Year Quota Projection (2027 - 2031)
  const ctxForecast = document.getElementById('chart-5year-forecast');
  if (ctxForecast && window.Chart) {
    new Chart(ctxForecast, {
      type: 'bar',
      data: {
        labels: ['2026-2027 (Hiện tại)', '2027-2028', '2028-2029', '2029-2030', '2030-2031'],
        datasets: [
          {
            label: 'Số lớp học',
            data: [14, 15, 15, 16, 16],
            type: 'line',
            borderColor: '#ef4444',
            backgroundColor: '#ef4444',
            borderWidth: 3,
            yAxisID: 'y1',
            tension: 0.2
          },
          {
            label: 'Biên chế viên chức đề xuất',
            data: [33, 35, 35, 37, 37],
            backgroundColor: '#0e7490',
            yAxisID: 'y'
          },
          {
            label: 'Hợp đồng lao động NĐ 111',
            data: [2, 2, 2, 2, 2],
            backgroundColor: '#f59e0b',
            yAxisID: 'y'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: false,
            min: 20,
            max: 45,
            title: { display: true, text: 'Tổng số người làm việc' }
          },
          y1: {
            position: 'right',
            min: 10,
            max: 20,
            grid: { drawOnChartArea: false },
            title: { display: true, text: 'Số lớp học' }
          }
        },
        plugins: {
          legend: {
            position: 'bottom',
            labels: { boxWidth: 12, font: { family: 'Inter', size: 12 } }
          }
        }
      }
    });
  }
}

// 8. DOWNLOAD HELPER FUNCTIONS
function downloadFile(type) {
  let fileUrl = '';
  let fileName = '';

  if (type === 'excel') {
    fileUrl = 'assets/Bo-Bieu-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.xlsx';
    fileName = 'Bo-Bieu-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.xlsx';
  } else if (type === 'word') {
    fileUrl = 'assets/De-An-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.docx';
    fileName = 'De-An-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.docx';
  }

  const link = document.createElement('a');
  link.href = fileUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
