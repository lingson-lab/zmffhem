(() => {
  "use strict";

  const BRAND_COLORS = [
    "#0300CE", "#5856D6", "#7D7AFF", "#A7A5FF", "#CAC9FF",
    "#34329B", "#7674B8", "#9A98C9"
  ];

  const TIME_SLOTS = [
    { label: "06:00~09:00", start: 6, end: 9 },
    { label: "09:00~12:00", start: 9, end: 12 },
    { label: "12:00~15:00", start: 12, end: 15 },
    { label: "15:00~18:00", start: 15, end: 18 },
    { label: "18:00~21:00", start: 18, end: 21 },
    { label: "21:00~06:00", start: 21, end: 6, overnight: true }
  ];

  const FIELD_ALIASES = {
    date: ["수강시작일", "수강 시작일", "거래일시", "거래 일시", "시작일자", "시작일"],
    facility: ["등록명", "등록 명", "시설명", "강좌명", "프로그램명", "상품명"],
    status: ["거래구분", "거래 구분", "거래상태", "거래 상태", "상태"],
    transactionType: ["거래종류", "거래 종류", "거래유형", "거래 유형"],
    amount: ["실매출액", "실 매출액", "실결제금액", "실 결제금액", "매출금액", "결제금액", "금액"],
    member: ["회원명", "회원 명", "성명", "입주민명"],
    building: ["동", "동명", "동 명"],
    room: ["호수", "호 수", "호"],
    channel: ["담당자", "접수채널", "접수 채널", "채널", "등록담당자"],
    count: ["건수", "수량"],
    accessTime: ["출입시간", "출입 시간", "시간"],
    accessDate: ["출입일자", "출입 일자", "일자", "출입날짜"]
  };

  const state = {
    rows: [],
    keys: {},
    months: [],
    accessRows: [],
    accessKeys: {}
  };

  const $ = (id) => document.getElementById(id);

  const els = {
    excelFile: $("excelFile"),
    accessFile: $("accessFile"),
    monthSelect: $("monthSelect"),
    printBtn: $("printBtn"),
    statusText: $("statusText"),
    statusDot: $("statusDot"),
    periodText: $("periodText"),

    totalSales: $("totalSales"),
    totalTransactions: $("totalTransactions"),
    uniqueMembers: $("uniqueMembers"),
    averagePayment: $("averagePayment"),

    salesDelta: $("salesDelta"),
    salesBreakdown: $("salesBreakdown"),
    transactionsDelta: $("transactionsDelta"),
    membersDelta: $("membersDelta"),
    averageDelta: $("averageDelta"),
    comparisonNote: $("comparisonNote"),

    highlightCardsSection: $("highlightCardsSection"),
    maxSalesName: $("maxSalesName"),
    maxSalesAmount: $("maxSalesAmount"),
    maxSalesDetail: $("maxSalesDetail"),
    minSalesName: $("minSalesName"),
    minSalesAmount: $("minSalesAmount"),
    minSalesDetail: $("minSalesDetail"),

    facilityAccordionBtn: $("facilityAccordionBtn"),
    facilityAccordionContent: $("facilityAccordionContent"),
    facilityMonthLabel: $("facilityMonthLabel"),
    facilityTable: $("facilityTable"),
    donutChart: $("donutChart"),
    donutTotal: $("donutTotal"),
    donutLegend: $("donutLegend"),

    peakTimeCardsSection: $("peakTimeCardsSection"),
    peakTimeLabel: $("peakTimeLabel"),
    peakTimeCount: $("peakTimeCount"),
    peakTimeDetail: $("peakTimeDetail"),

    accessSection: $("accessSection"),
    accessAccordionBtn: $("accessAccordionBtn"),
    accessAccordionContent: $("accessAccordionContent"),
    accessMonthLabel: $("accessMonthLabel"),
    accessTable: $("accessTable"),
    accessChart: $("accessChart"),

    generatedAt: $("generatedAt")
  };

  function normalize(value) {
    return String(value ?? "")
      .trim()
      .replace(/\s+/g, "")
      .replace(/[()_\-./]/g, "")
      .toLowerCase();
  }

  function findHeader(headers, aliases) {
    const normalized = headers.map(h => ({ raw: h, norm: normalize(h) }));

    for (const alias of aliases) {
      const target = normalize(alias);
      const exact = normalized.find(item => item.norm === target);
      if (exact) return exact.raw;
    }

    for (const alias of aliases) {
      const target = normalize(alias);
      const partial = normalized.find(item =>
        item.norm.includes(target) || target.includes(item.norm)
      );
      if (partial) return partial.raw;
    }

    return null;
  }

  function detectKeys(headers, isAccess = false) {
    const keys = {};
    const aliases = isAccess
      ? { accessTime: FIELD_ALIASES.accessTime, accessDate: FIELD_ALIASES.accessDate, member: FIELD_ALIASES.member, building: FIELD_ALIASES.building }
      : FIELD_ALIASES;

    for (const [name, aliasList] of Object.entries(aliases)) {
      keys[name] = findHeader(headers, aliasList);
    }

    if (isAccess) {
      state.accessKeys = keys;
    } else {
      state.keys = keys;
    }

    if (!isAccess) {
      const required = ["date", "facility", "amount", "transactionType"];
      const missing = required.filter(k => !keys[k]);

      if (missing.length) {
        throw new Error(
          "필수 헤더를 찾지 못했습니다: " +
          missing.map(k => ({
            date: "수강시작일",
            facility: "등록명",
            amount: "실매출액",
            transactionType: "거래종류"
          }[k])).join(", ")
        );
      }
    }
  }

  function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;

    const text = String(value ?? "")
      .replace(/₩|원|,/g, "")
      .replace(/\s/g, "")
      .replace(/[^\d.+-]/g, "");

    const n = Number(text);
    return Number.isFinite(n) ? n : 0;
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

    if (typeof value === "number" && window.XLSX?.SSF?.parse_date_code) {
      const d = XLSX.SSF.parse_date_code(value);

      if (d) {
        return new Date(
          d.y,
          d.m - 1,
          d.d,
          d.H || 0,
          d.M || 0,
          Math.floor(d.S || 0)
        );
      }
    }

    const text = String(value ?? "").trim();
    const m = text.match(/^(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);

    if (m) {
      return new Date(
        Number(m[1]),
        Number(m[2]) - 1,
        Number(m[3])
      );
    }

    const fallback = new Date(text);
    return Number.isNaN(fallback.getTime()) ? null : fallback;
  }

  function parseTime(value) {
    const text = String(value ?? "").trim();
    const m = text.match(/(\d{1,2}):(\d{2})/);
    if (m) {
      return { hour: Number(m[1]), minute: Number(m[2]) };
    }
    return null;
  }

  function monthKey(date) {
    if (!date) return null;

    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  function previousMonthKey(key) {
    if (!key) return null;

    const [year, month] = key.split("-").map(Number);
    const d = new Date(year, month - 2, 1);

    return monthKey(d);
  }

  function formatMonth(key) {
    if (!key) return "-";

    const [y, m] = key.split("-");
    return `${y}.${m}`;
  }

  function formatWon(n) {
    return `${Math.round(n).toLocaleString("ko-KR")}원`;
  }

  function formatDate(d) {
    return `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, "0")}.${String(d.getDate()).padStart(2, "0")}`;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function setStatus(message, type = "") {
    els.statusText.textContent = message;
    els.statusDot.className = `status-dot ${type}`.trim();
  }

  function normalizeFacility(value) {
    const text = String(value ?? "").trim();

    if (!text) return "기타";

    const first = text.split(/[▶>›]/)[0].trim();
    return first || text;
  }

  function countOf(row) {
    return state.keys.count
      ? Math.max(1, parseNumber(row[state.keys.count]))
      : 1;
  }

  function classifyTransaction(row) {
    const tx = state.keys.transactionType
      ? String(row[state.keys.transactionType] ?? "").trim()
      : "";

    if (/환불/.test(tx)) return "refund";
    if (/취소/.test(tx)) return "cancel";
    if (/매출/.test(tx)) return "sale";

    return "other";
  }

  function netAmount(row) {
    const type = classifyTransaction(row);
    const amount = Math.abs(parseNumber(row[state.keys.amount]));

    if (type === "sale") return amount;
    if (type === "refund") return -amount;

    return 0;
  }

  function isTestData(row) {
    const building = String(row[state.keys.building] ?? "").trim();
    const room = state.keys.room ? String(row[state.keys.room] ?? "").trim() : "";
    
    if (building === "999" || building === "9999") return true;
    if (room === "9999" || room === "99999") return true;
    
    return false;
  }

  function prepareRows(rows) {
    return rows
      .map(row => {
        const d = parseDate(row[state.keys.date]);

        const type = classifyTransaction(row);

        return {
          ...row,
          __date: d,
          __month: monthKey(d),
          __facility: normalizeFacility(row[state.keys.facility]),
          __amount: parseNumber(row[state.keys.amount]),
          __netAmount: netAmount(row),
          __transactionType: type,
          __count: countOf(row)
        };
      })
      .filter(row => row.__date && !isTestData(row));
  }

  function prepareAccessRows(rows) {
    return rows
      .map(row => {
        const d = parseDate(row[state.accessKeys.accessDate]);
        const t = parseTime(row[state.accessKeys.accessTime]);

        if (!d || !t) return null;

        return {
          ...row,
          __date: d,
          __month: monthKey(d),
          __hour: t.hour,
          __minute: t.minute,
          __member: state.accessKeys.member ? String(row[state.accessKeys.member] ?? "").trim() : null
        };
      })
      .filter(Boolean);
  }

  function getTimeSlot(hour) {
    for (const slot of TIME_SLOTS) {
      if (slot.overnight) {
        if (hour >= slot.start || hour < slot.end) return slot;
      } else {
        if (hour >= slot.start && hour < slot.end) return slot;
      }
    }
    return TIME_SLOTS[0];
  }

  function percent(n, d) {
    return d ? (n / d) * 100 : 0;
  }

  function getSalesBreakdown(rows) {
    let grossSales = 0;
    let refunds = 0;
    let refundCount = 0;

    rows.forEach(row => {
      if (row.__transactionType === "sale") {
        grossSales += Math.abs(row.__amount);
      } else if (row.__transactionType === "refund") {
        refunds += Math.abs(row.__amount);
        refundCount += row.__count;
      }
    });

    return {
      grossSales,
      refunds,
      refundCount,
      netSales: grossSales - refunds
    };
  }

  function getMetrics(rows) {
    const salesRows = rows.filter(row => row.__transactionType === "sale");
    const netSales = rows.reduce((sum, row) => sum + row.__netAmount, 0);

    const transactions = salesRows.reduce(
      (sum, row) => sum + row.__count,
      0
    );

    const members = state.keys.member
      ? new Set(
          salesRows
            .map(row => row[state.keys.member])
            .filter(m => m)
        ).size
      : null;

    const average = transactions ? netSales / transactions : 0;

    return { sales: netSales, transactions, members, average };
  }

  function getComparisonRows(month, currentRows) {
    const prev = previousMonthKey(month);

    if (!prev) {
      return {
        key: null,
        rows: [],
        available: false,
        label: `기준월 ${formatMonth(month)}`
      };
    }

    const prevRows = state.rows.filter(row => row.__month === prev);

    if (!prevRows.length) {
      return {
        key: prev,
        rows: [],
        available: false,
        label: `기준월 ${formatMonth(month)} · 전월 ${formatMonth(prev)}(데이터 없음)`
      };
    }

    return {
      key: prev,
      rows: prevRows,
      available: true,
      label: `기준월 ${formatMonth(month)} · 전월 ${formatMonth(prev)}`
    };
  }

  function deltaInfo(current, previous, available) {
    if (!available || previous === null || typeof previous !== "number") {
      return { text: "-", cls: "neutral" };
    }

    if (previous === 0) {
      if (current === 0) {
        return { text: "→", cls: "flat" };
      }

      return { text: "신규", cls: "up" };
    }

    const delta = ((current - previous) / previous) * 100;

    if (Math.abs(delta) < 0.5) {
      return { text: "→", cls: "flat" };
    }

    if (delta > 0) {
      return {
        text: `↑ ${Math.abs(delta).toFixed(1)}%`,
        cls: "up"
      };
    }

    return {
      text: `↓ ${Math.abs(delta).toFixed(1)}%`,
      cls: "down"
    };
  }

  function shortDeltaInfo(current, previous, available) {
    return deltaInfo(current, previous, available);
  }

  function setDeltaElement(el, deltaInfo) {
    el.className = `kpi-delta ${deltaInfo.cls}`;
    el.textContent = deltaInfo.text;
  }

  function toggleAccordion(btn, content) {
    btn.classList.toggle("open");
    content.classList.toggle("collapsed");
  }

  function renderFacility(rows, comparisonRows, comparisonAvailable) {
    const currentMap = new Map();
    const previousMap = new Map();

    let totalSales = 0;
    let previousTotalSales = 0;
    let totalTransactions = 0;
    const allMembers = state.keys.member ? new Set() : null;

    rows.forEach(row => {
      if (row.__netAmount === 0) return;

      const name = row.__facility;

      if (!currentMap.has(name)) {
        currentMap.set(name, {
          name,
          sales: 0,
          transactions: 0,
          members: state.keys.member ? new Set() : null
        });
      }

      const item = currentMap.get(name);
      item.sales += row.__netAmount;
      item.transactions += row.__count;

      if (state.keys.member && row[state.keys.member]) {
        item.members.add(row[state.keys.member]);
        allMembers.add(row[state.keys.member]);
      }

      totalSales += row.__netAmount;
      totalTransactions += row.__count;
    });

    comparisonRows.forEach(row => {
      if (row.__netAmount === 0) return;

      const name = row.__facility;

      if (!previousMap.has(name)) {
        previousMap.set(name, {
          name,
          sales: 0,
          transactions: 0,
          members: state.keys.member ? new Set() : null
        });
      }

      const item = previousMap.get(name);
      item.sales += row.__netAmount;
      item.transactions += row.__count;

      if (state.keys.member && row[state.keys.member]) {
        item.members.add(row[state.keys.member]);
      }

      previousTotalSales += row.__netAmount;
    });

    const items = [...currentMap.values()]
      .map(item => ({
        ...item,
        prevSales: previousMap.get(item.name)?.sales ?? 0
      }))
      .sort((a, b) => b.sales - a.sales)
      .filter(item => item.sales !== 0);

    if (items.length > 0) {
      const maxItem = items[0];
      const minItem = items[items.length - 1];

      els.highlightCardsSection.style.display = "grid";
      
      els.maxSalesName.textContent = escapeHtml(maxItem.name);
      els.maxSalesAmount.textContent = formatWon(maxItem.sales);
      els.maxSalesDetail.textContent = `거래 ${maxItem.transactions.toLocaleString("ko-KR")}건 · 회원 ${state.keys.member ? maxItem.members.size.toLocaleString("ko-KR") : "-"}명`;

      els.minSalesName.textContent = escapeHtml(minItem.name);
      els.minSalesAmount.textContent = formatWon(minItem.sales);
      els.minSalesDetail.textContent = `거래 ${minItem.transactions.toLocaleString("ko-KR")}건 · 회원 ${state.keys.member ? minItem.members.size.toLocaleString("ko-KR") : "-"}명`;
    } else {
      els.highlightCardsSection.style.display = "none";
    }

    if (!items.length) {
      els.facilityTable.innerHTML =
        '<tr><td colspan="7" class="empty">데이터 없음</td></tr>';

      els.donutChart.style.background =
        "conic-gradient(#E9ECF3 0 100%)";

      els.donutLegend.innerHTML = "";
      els.donutTotal.textContent = "0원";
      return;
    }

    els.facilityTable.innerHTML =
      items.map(item => {
        const delta = shortDeltaInfo(
          item.sales,
          item.prevSales,
          comparisonAvailable
        );

        return `
          <tr>
            <td><strong>${escapeHtml(item.name)}</strong></td>
            <td class="num">${formatWon(item.sales)}</td>
            <td class="num prev-sales">${comparisonAvailable ? formatWon(item.prevSales) : "-"}</td>
            <td class="num"><span class="mom-value ${delta.cls}">${delta.text}</span></td>
            <td class="num">${item.transactions.toLocaleString("ko-KR")}건</td>
            <td class="num">${state.keys.member ? (item.members ? item.members.size.toLocaleString("ko-KR") + "명" : "-") : "-"}</td>
            <td class="num">${totalSales ? percent(item.sales, totalSales).toFixed(1) : "0.0"}%</td>
          </tr>
        `;
      }).join("") +
      `
        <tr class="total-row">
          <td>합계</td>
          <td class="num">${formatWon(totalSales)}</td>
          <td class="num prev-sales">${comparisonAvailable ? formatWon(previousTotalSales) : "-"}</td>
          <td class="num">
            <span class="mom-value ${shortDeltaInfo(totalSales, previousTotalSales, comparisonAvailable).cls}">
              ${shortDeltaInfo(totalSales, previousTotalSales, comparisonAvailable).text}
            </span>
          </td>
          <td class="num">${totalTransactions.toLocaleString("ko-KR")}건</td>
          <td class="num">${state.keys.member ? allMembers.size.toLocaleString("ko-KR") + "명" : "-"}</td>
          <td class="num">100.0%</td>
        </tr>
      `;

    const donutItems = [...currentMap.values()]
      .filter(item => item.sales > 0)
      .sort((a, b) => b.sales - a.sales);

    let cursor = 0;
    const segments = [];

    donutItems.slice(0, 8).forEach((item, i) => {
      const p = percent(item.sales, totalSales);
      const next = cursor + p;

      segments.push(
        `${BRAND_COLORS[i % BRAND_COLORS.length]} ${cursor}% ${next}%`
      );

      cursor = next;
    });

    if (cursor < 100) {
      segments.push(`#E9ECF3 ${cursor}% 100%`);
    }

    els.donutChart.style.background = segments.length
      ? `conic-gradient(${segments.join(",")})`
      : "conic-gradient(#E9ECF3 0 100%)";

    els.donutTotal.textContent = formatWon(totalSales);

    els.donutLegend.innerHTML = donutItems
      .slice(0, 6)
      .map((item, i) => `
        <div class="legend-row">
          <i class="legend-dot" style="background:${BRAND_COLORS[i % BRAND_COLORS.length]}"></i>
          <strong>${escapeHtml(item.name)}</strong>
          <span>${percent(item.sales, totalSales).toFixed(1)}%</span>
        </div>
      `)
      .join("");
  }

  function renderAccess(rows) {
    if (!rows.length) {
      els.accessTable.innerHTML =
        '<tr><td colspan="4" class="empty">데이터 없음</td></tr>';
      els.peakTimeCardsSection.style.display = "none";
      return;
    }

    const slots = new Map();
    TIME_SLOTS.forEach(slot => {
      slots.set(slot.label, {
        count: 0,
        members: new Set()
      });
    });

    rows.forEach(row => {
      const slot = getTimeSlot(row.__hour);
      const data = slots.get(slot.label);
      data.count += 1;
      if (row.__member) {
        data.members.add(row.__member);
      }
    });

    const totalAccess = rows.length;
    const slotData = TIME_SLOTS.map(slot => ({
      label: slot.label,
      count: slots.get(slot.label).count,
      members: slots.get(slot.label).members.size
    }));

    // 피크 시간대 찾기
    const peakTimeData = slotData.reduce((max, current) => 
      current.count > max.count ? current : max
    );

    // 피크 시간대 카드 표시
    if (peakTimeData.count > 0) {
      els.peakTimeCardsSection.style.display = "grid";
      els.peakTimeLabel.textContent = peakTimeData.label;
      els.peakTimeCount.textContent = `${peakTimeData.count.toLocaleString("ko-KR")}회`;
      els.peakTimeDetail.textContent = `이용자 ${peakTimeData.members.toLocaleString("ko-KR")}명`;
    }

    els.accessTable.innerHTML = slotData
      .map(item => `
        <tr>
          <td><strong>${item.label}</strong></td>
          <td class="num">${item.count.toLocaleString("ko-KR")}회</td>
          <td class="num">${item.members.toLocaleString("ko-KR")}명</td>
          <td class="num">${totalAccess ? percent(item.count, totalAccess).toFixed(1) : "0.0"}%</td>
        </tr>
      `)
      .join("") +
      `
        <tr class="total-row">
          <td>합계</td>
          <td class="num">${totalAccess.toLocaleString("ko-KR")}회</td>
          <td class="num">${new Set(rows.map(r => r.__member)).size.toLocaleString("ko-KR")}명</td>
          <td class="num">100.0%</td>
        </tr>
      `;

    drawAccessChart(slotData, totalAccess);
  }

  function drawAccessChart(slotData, total) {
    const canvas = els.accessChart;
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    
    const leftPadding = 50;
    const rightPadding = 20;
    const topPadding = 30;
    const bottomPadding = 80;
    
    const width = canvas.width - leftPadding - rightPadding;
    const height = canvas.height - topPadding - bottomPadding;
    const barWidth = width / slotData.length * 0.65;
    const barGap = width / slotData.length;

    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = "#6F7787";
    ctx.font = "12px Pretendard";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    for (let i = 0; i <= 5; i++) {
      const y = topPadding + height - (height * i / 5);
      ctx.fillText(Math.round(total * i / 5), leftPadding - 12, y);
    }

    ctx.strokeStyle = "#E7EAF0";
    ctx.lineWidth = 1;
    for (let i = 0; i <= 5; i++) {
      const y = topPadding + height - (height * i / 5);
      ctx.beginPath();
      ctx.moveTo(leftPadding, y);
      ctx.lineTo(canvas.width - rightPadding, y);
      ctx.stroke();
    }

    slotData.forEach((item, idx) => {
      const x = leftPadding + idx * barGap + (barGap - barWidth) / 2;
      const barHeight = (item.count / total) * height;
      const y = topPadding + height - barHeight;

      ctx.fillStyle = "#0300CE";
      ctx.fillRect(x, y, barWidth, barHeight);
    });

    ctx.fillStyle = "#6F7787";
    ctx.font = "12px Pretendard";
    ctx.textAlign = "right";
    ctx.textBaseline = "top";

    slotData.forEach((item, idx) => {
      const x = leftPadding + idx * barGap + (barGap - barWidth) / 2 + barWidth / 2;
      const y = topPadding + height + 15;

      ctx.save();
      ctx.translate(x, y);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(item.label, 0, 0);
      ctx.restore();
    });
  }

  function render(month) {
    const rows = state.rows.filter(row => row.__month === month);
    const comparison = getComparisonRows(month, rows);

    const currentMetrics = getMetrics(rows);
    const previousMetrics = getMetrics(comparison.rows);
    const salesBreakdown = getSalesBreakdown(rows);

    els.totalSales.textContent = formatWon(currentMetrics.sales);

    if (els.salesBreakdown) {
      els.salesBreakdown.textContent =
        `매출 ${formatWon(salesBreakdown.grossSales)} − 환불 ${formatWon(salesBreakdown.refunds)} (${salesBreakdown.refundCount.toLocaleString("ko-KR")}건)`;
    }
    els.totalTransactions.textContent =
      `${currentMetrics.transactions.toLocaleString("ko-KR")}건`;

    els.uniqueMembers.textContent =
      currentMetrics.members === null
        ? "-"
        : `${currentMetrics.members.toLocaleString("ko-KR")}명`;

    els.averagePayment.textContent =
      formatWon(currentMetrics.average);

    setDeltaElement(
      els.salesDelta,
      deltaInfo(
        currentMetrics.sales,
        previousMetrics.sales,
        comparison.available
      )
    );

    setDeltaElement(
      els.transactionsDelta,
      deltaInfo(
        currentMetrics.transactions,
        previousMetrics.transactions,
        comparison.available
      )
    );

    setDeltaElement(
      els.membersDelta,
      state.keys.member
        ? deltaInfo(
            currentMetrics.members,
            previousMetrics.members,
            comparison.available
          )
        : { text: "회원명 컬럼 없음", cls: "neutral" }
    );

    setDeltaElement(
      els.averageDelta,
      deltaInfo(
        currentMetrics.average,
        previousMetrics.average,
        comparison.available
      )
    );

    els.comparisonNote.textContent = comparison.available
      ? comparison.label
      : `전월 비교 불가 · ${comparison.label}`;

    const dates = rows
      .map(row => row.__date)
      .filter(Boolean)
      .sort((a, b) => a - b);

    if (dates.length) {
      els.periodText.textContent =
        `${formatDate(dates[0])} ~ ${formatDate(dates[dates.length - 1])} · 선택 기준월 ${formatMonth(month)}`;
    }

    els.facilityMonthLabel.textContent = comparison.available
      ? `${formatMonth(month)} vs ${formatMonth(comparison.key)}`
      : formatMonth(month);

    renderFacility(
      rows,
      comparison.rows,
      comparison.available
    );

    if (state.accessRows.length > 0) {
      els.accessSection.style.display = "grid";
      const accessRows = state.accessRows.filter(row => row.__month === month);
      els.accessMonthLabel.textContent = formatMonth(month);
      renderAccess(accessRows);
    }

    els.generatedAt.textContent =
      `생성 ${new Intl.DateTimeFormat("ko-KR", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit"
      }).format(new Date())}`;
  }

  function populateMonths() {
    state.months = [
      ...new Set(
        state.rows
          .map(row => row.__month)
          .filter(Boolean)
      )
    ].sort().reverse();

    els.monthSelect.innerHTML = "";

    if (!state.months.length) {
      els.monthSelect.disabled = true;
      els.monthSelect.innerHTML =
        "<option>기준월 없음</option>";
      return;
    }

    state.months.forEach(month => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = formatMonth(month);
      els.monthSelect.appendChild(option);
    });

    els.monthSelect.disabled = false;
    els.monthSelect.value = state.months[0];

    render(state.months[0]);
  }

  async function loadFile(file, isAccess = false) {
    try {
      if (isAccess) {
        setStatus(`${file.name} 출입 데이터를 읽는 중입니다.`);
      } else {
        setStatus(`${file.name} 파일을 읽는 중입니다.`);
      }

      const buffer = await file.arrayBuffer();

      const wb = XLSX.read(buffer, {
        type: "array",
        cellDates: true
      });

      const sheet = wb.Sheets[wb.SheetNames[0]];

      if (!sheet) {
        throw new Error("첫 번째 시트를 찾을 수 없습니다.");
      }

      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
        raw: true
      });

      if (!rows.length) {
        throw new Error("분석할 데이터가 없습니다.");
      }

      if (isAccess) {
        detectKeys(Object.keys(rows[0]), true);
        state.accessRows = prepareAccessRows(rows);
        
        if (state.rows.length > 0) {
          render(state.months[0]);
          setStatus(
            `${file.name} · ${state.accessRows.length.toLocaleString("ko-KR")}행 분석 완료`,
            "ok"
          );
        }
      } else {
        detectKeys(Object.keys(rows[0]));
        state.rows = prepareRows(rows);

        const unknownTypes = [
          ...new Set(
            state.rows
              .filter(row => row.__transactionType === "other")
              .map(row => String(row[state.keys.transactionType] ?? "").trim())
              .filter(Boolean)
          )
        ];

        const testDataCount = rows.length - state.rows.length;

        populateMonths();

        if (unknownTypes.length) {
          setStatus(
            `${file.name} · ${state.rows.length.toLocaleString("ko-KR")}행 분석 완료${testDataCount > 0 ? ` (테스트 데이터 ${testDataCount}건 제외)` : ""} · 미인식 거래종류는 실매출에서 제외: ${unknownTypes.join(", ")}`,
            "warn"
          );
        } else {
          setStatus(
            `${file.name} · ${state.rows.length.toLocaleString("ko-KR")}행 분석 완료${testDataCount > 0 ? ` (테스트 데이터 ${testDataCount}건 제외)` : ""} · 실매출 = 매출 − 환불, 취소 제외`,
            "ok"
          );
        }
      }
    } catch (err) {
      console.error(err);

      setStatus(
        `분석 실패: ${err.message}`,
        "error"
      );
    }
  }

  els.excelFile.addEventListener("change", event => {
    const file = event.target.files?.[0];

    if (file) {
      loadFile(file, false);
    }
  });

  els.accessFile.addEventListener("change", event => {
    const file = event.target.files?.[0];

    if (file) {
      loadFile(file, true);
    }
  });

  els.monthSelect.addEventListener("change", event => {
    render(event.target.value);
  });

  if (els.facilityAccordionBtn && els.facilityAccordionContent) {
    els.facilityAccordionBtn.addEventListener("click", () => {
      toggleAccordion(els.facilityAccordionBtn, els.facilityAccordionContent);
    });
  }

  if (els.accessAccordionBtn && els.accessAccordionContent) {
    els.accessAccordionBtn.addEventListener("click", () => {
      toggleAccordion(els.accessAccordionBtn, els.accessAccordionContent);
    });
  }

  els.printBtn.addEventListener("click", () => {
    if (els.facilityAccordionContent.classList.contains("collapsed")) {
      toggleAccordion(els.facilityAccordionBtn, els.facilityAccordionContent);
    }
    if (els.accessAccordionContent && els.accessAccordionContent.classList.contains("collapsed")) {
      toggleAccordion(els.accessAccordionBtn, els.accessAccordionContent);
    }
    
    setTimeout(() => {
      window.print();
    }, 100);
  });

  if (!window.XLSX) {
    setStatus(
      "엑셀 분석 라이브러리를 불러오지 못했습니다.",
      "error"
    );
  }
})();
