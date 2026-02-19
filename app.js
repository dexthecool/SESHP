(() => {
  "use strict";

  const KEY_SEP = "\u001f";
  const MONTH_NAMES = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  const state = {
    files: [],
    processing: false,
  };

  const dom = {
    dropZone: document.getElementById("dropZone"),
    fileInput: document.getElementById("fileInput"),
    chooseFilesBtn: document.getElementById("chooseFilesBtn"),
    clearFilesBtn: document.getElementById("clearFilesBtn"),
    fileSummary: document.getElementById("fileSummary"),
    fileList: document.getElementById("fileList"),
    processBtn: document.getElementById("processBtn"),
    statusText: document.getElementById("statusText"),
    statusLog: document.getElementById("statusLog"),
  };

  init();

  function init() {
    bindFileInputEvents();
    bindDropEvents();
    bindActionEvents();
    updateModeCardStates();
    updateProcessState();
    renderSelectedFiles();
  }

  function bindFileInputEvents() {
    dom.chooseFilesBtn.addEventListener("click", () => dom.fileInput.click());
    dom.fileInput.addEventListener("change", (event) => {
      const files = Array.from(event.target.files || []);
      addFiles(files);
      dom.fileInput.value = "";
    });
    dom.clearFilesBtn.addEventListener("click", clearFiles);
  }

  function bindDropEvents() {
    const activate = () => dom.dropZone.classList.add("is-active");
    const deactivate = () => dom.dropZone.classList.remove("is-active");

    dom.dropZone.addEventListener("dragenter", (event) => {
      event.preventDefault();
      activate();
    });
    dom.dropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      activate();
    });
    dom.dropZone.addEventListener("dragleave", (event) => {
      event.preventDefault();
      deactivate();
    });
    dom.dropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      deactivate();
      const files = Array.from(event.dataTransfer?.files || []);
      addFiles(files);
    });

    dom.dropZone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        dom.fileInput.click();
      }
    });
  }

  function bindActionEvents() {
    dom.processBtn.addEventListener("click", handleProcess);
    document.querySelectorAll("input[name='outputMode']").forEach((radio) => {
      radio.addEventListener("change", () => {
        updateModeCardStates();
        updateProcessState();
      });
    });
  }

  function updateModeCardStates() {
    const cards = document.querySelectorAll(".mode-card");
    cards.forEach((card) => {
      const input = card.querySelector("input[name='outputMode']");
      card.classList.toggle("is-selected", Boolean(input?.checked));
    });
  }

  function addFiles(incomingFiles) {
    if (!incomingFiles.length) {
      return;
    }

    const existing = new Set(state.files.map(fileUniqueKey));
    let added = 0;
    let skipped = 0;

    for (const file of incomingFiles) {
      if (!isSupportedInputFile(file.name)) {
        skipped += 1;
        continue;
      }

      const key = fileUniqueKey(file);
      if (existing.has(key)) {
        skipped += 1;
        continue;
      }

      state.files.push(file);
      existing.add(key);
      added += 1;
    }

    renderSelectedFiles();
    updateProcessState();

    if (added > 0) {
      addStatus(`${added} file(s) added.`, "ok");
    }
    if (skipped > 0) {
      addStatus(`${skipped} file(s) skipped (duplicate or unsupported).`, "warn");
    }
  }

  function clearFiles() {
    state.files = [];
    renderSelectedFiles();
    updateProcessState();
    addStatus("Selected files cleared.");
  }

  function renderSelectedFiles() {
    if (!state.files.length) {
      dom.fileSummary.textContent = "No files selected.";
      dom.fileList.innerHTML = "";
      return;
    }

    const zipCount = state.files.filter((file) => file.name.toLowerCase().endsWith(".zip")).length;
    const jsonCount = state.files.filter((file) => file.name.toLowerCase().endsWith(".json")).length;
    dom.fileSummary.textContent = `${state.files.length} file(s) selected (${jsonCount} JSON, ${zipCount} ZIP).`;

    dom.fileList.innerHTML = state.files
      .map(
        (file) =>
          `<li><span class="name" title="${escapeHtml(file.name)}">${escapeHtml(
            file.name
          )}</span><span class="size">${formatFileSize(file.size)}</span></li>`
      )
      .join("");
  }

  function updateProcessState() {
    const mode = getSelectedMode();
    if (state.processing) {
      dom.statusText.textContent = "Processing your data...";
      dom.processBtn.textContent = "Processing...";
      dom.processBtn.disabled = true;
      return;
    }

    const hasFiles = state.files.length > 0;
    dom.processBtn.disabled = !hasFiles;
    dom.processBtn.textContent =
      mode === "pdf" ? "Process and Download PDF" : "Process and Download Excel";
    dom.statusText.textContent = hasFiles
      ? "Ready to process."
      : "Add files to enable processing.";
  }

  async function handleProcess() {
    if (!state.files.length || state.processing) {
      return;
    }

    state.processing = true;
    resetStatusLog();
    updateProcessState();

    try {
      validateLibraries(getSelectedMode());
      addStatus("Reading selected files...");
      const parsedInput = await parseUploadedFiles(state.files);

      if (!parsedInput.records.length) {
        throw new Error("No valid JSON streaming records were found in the selected files.");
      }

      addStatus(
        `Loaded ${formatNumber(parsedInput.records.length)} total records from ${parsedInput.parsedJsonCount} JSON file(s).`,
        "ok"
      );

      const report = buildReport(parsedInput.records, parsedInput);
      if (report.songRecordsUsed === 0) {
        throw new Error("No song records were found after filtering out non-song entries.");
      }

      const mode = getSelectedMode();
      if (mode === "excel") {
        const filename = exportExcelReport(report);
        addStatus(`Excel report generated: ${filename}`, "ok");
      } else {
        const filename = await exportSimplifiedPdf(report);
        addStatus(`PDF report generated: ${filename}`, "ok");
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      addStatus(message, "error");
    } finally {
      state.processing = false;
      updateProcessState();
    }
  }

  function validateLibraries(mode) {
    if (typeof JSZip === "undefined") {
      throw new Error("JSZip failed to load.");
    }
    if (mode === "excel" && typeof XLSX === "undefined") {
      throw new Error("XLSX library failed to load.");
    }
    if (mode === "pdf") {
      if (!window.jspdf?.jsPDF) {
        throw new Error("jsPDF failed to load.");
      }
      if (typeof Chart === "undefined") {
        throw new Error("Chart.js failed to load.");
      }
      if (typeof window.jspdf.jsPDF.API.autoTable !== "function") {
        throw new Error("jsPDF AutoTable plugin failed to load.");
      }
    }
  }

  async function parseUploadedFiles(files) {
    const records = [];
    let parsedJsonCount = 0;
    let zipCount = 0;
    let jsonCount = 0;
    let skippedCount = 0;

    for (const file of files) {
      const lower = file.name.toLowerCase();
      if (lower.endsWith(".zip")) {
        zipCount += 1;
        addStatus(`Scanning ZIP: ${file.name}`);
        const zipRecords = await parseZipFile(file);
        records.push(...zipRecords.records);
        parsedJsonCount += zipRecords.parsedJsonCount;
        skippedCount += zipRecords.skippedEntries;
      } else if (lower.endsWith(".json")) {
        jsonCount += 1;
        try {
          const jsonRecords = await parseJsonBlob(file, file.name);
          records.push(...jsonRecords);
          parsedJsonCount += 1;
        } catch (error) {
          skippedCount += 1;
          addStatus(`Failed to parse ${file.name}: ${errorMessage(error)}`, "warn");
        }
      } else {
        skippedCount += 1;
        addStatus(`Unsupported file skipped: ${file.name}`, "warn");
      }
    }

    if (zipCount > 0) {
      addStatus(`ZIP files processed: ${zipCount}`);
    }
    if (jsonCount > 0) {
      addStatus(`Direct JSON files processed: ${jsonCount}`);
    }
    if (skippedCount > 0) {
      addStatus(`Skipped entries/files: ${skippedCount}`, "warn");
    }

    return { records, parsedJsonCount, zipCount, jsonCount, skippedCount };
  }

  async function parseZipFile(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const entries = Object.values(zip.files).filter(
      (entry) => !entry.dir && entry.name.toLowerCase().endsWith(".json")
    );

    if (!entries.length) {
      addStatus(`No JSON files found inside ${file.name}.`, "warn");
      return { records: [], parsedJsonCount: 0, skippedEntries: 1 };
    }

    const records = [];
    let parsedJsonCount = 0;
    let skippedEntries = 0;

    for (const entry of entries) {
      try {
        const text = await entry.async("string");
        const entryRecords = parseJsonText(text, `${file.name} -> ${entry.name}`);
        records.push(...entryRecords);
        parsedJsonCount += 1;
      } catch (error) {
        skippedEntries += 1;
        addStatus(`Failed to parse ${entry.name}: ${errorMessage(error)}`, "warn");
      }
    }

    addStatus(`Parsed ${parsedJsonCount} JSON file(s) from ${file.name}.`);
    return { records, parsedJsonCount, skippedEntries };
  }

  async function parseJsonBlob(file, label) {
    const text = await file.text();
    return parseJsonText(text, label);
  }

  function parseJsonText(text, label) {
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (error) {
      throw new Error(`Invalid JSON in ${label}`);
    }

    if (!Array.isArray(parsed)) {
      throw new Error(`Expected a JSON array in ${label}`);
    }

    return parsed.filter((item) => item && typeof item === "object" && !Array.isArray(item));
  }

  function buildReport(records, parsedInputMeta) {
    const songMap = new Map();
    const artistMap = new Map();
    const monthlySongMap = new Map();
    const monthlyMs = new Map();
    const yearlyMs = new Map();

    let ignoredNonSong = 0;
    let ignoredBadTimestamp = 0;
    let songRecordsUsed = 0;

    for (const record of records) {
      if (!isSongRecord(record)) {
        ignoredNonSong += 1;
        continue;
      }

      const timestamp = parseTimestamp(record.ts);
      if (!timestamp) {
        ignoredBadTimestamp += 1;
        continue;
      }

      const song = cleanString(record.master_metadata_track_name);
      const artist = cleanString(record.master_metadata_album_artist_name);
      const album = cleanString(record.master_metadata_album_album_name);
      const uri = cleanString(record.spotify_track_uri);
      const msPlayed = sanitizeMs(record.ms_played);
      const skipped = Boolean(record.skipped);
      const songKey = makeSongKey(song, artist);
      const monthKey = timestampToMonthKey(timestamp);
      const yearKey = timestamp.getUTCFullYear();

      const songEntry = songMap.get(songKey) || {
        song,
        artist,
        playCount: 0,
        skipCount: 0,
        totalMs: 0,
        albumCounter: new Map(),
        uriCounter: new Map(),
      };
      songEntry.playCount += 1;
      songEntry.skipCount += skipped ? 1 : 0;
      songEntry.totalMs += msPlayed;
      incrementCounter(songEntry.albumCounter, album);
      incrementCounter(songEntry.uriCounter, uri);
      songMap.set(songKey, songEntry);

      const artistEntry = artistMap.get(artist) || {
        artist,
        playCount: 0,
        skipCount: 0,
        totalMs: 0,
      };
      artistEntry.playCount += 1;
      artistEntry.skipCount += skipped ? 1 : 0;
      artistEntry.totalMs += msPlayed;
      artistMap.set(artist, artistEntry);

      const monthSongs = monthlySongMap.get(monthKey) || new Map();
      const monthSongEntry = monthSongs.get(songKey) || { playCount: 0, totalMs: 0 };
      monthSongEntry.playCount += 1;
      monthSongEntry.totalMs += msPlayed;
      monthSongs.set(songKey, monthSongEntry);
      monthlySongMap.set(monthKey, monthSongs);

      monthlyMs.set(monthKey, (monthlyMs.get(monthKey) || 0) + msPlayed);
      yearlyMs.set(yearKey, (yearlyMs.get(yearKey) || 0) + msPlayed);

      songRecordsUsed += 1;
    }

    const songRows = Array.from(songMap.values()).map((entry) => {
      const album = mostCommonCounterKey(entry.albumCounter);
      const uri = mostCommonCounterKey(entry.uriCounter);
      const totalMinutes = entry.totalMs / 60000;
      return {
        song: entry.song,
        artist: entry.artist,
        album,
        uri,
        playCount: entry.playCount,
        skipCount: entry.skipCount,
        skipRate: entry.playCount ? entry.skipCount / entry.playCount : 0,
        totalMs: entry.totalMs,
        totalMinutes,
        totalHours: totalMinutes / 60,
      };
    });

    songRows.sort((a, b) => compareSongsByPlayCount(a, b));

    const top10ByPlayCount = [...songRows].sort(compareSongsByPlayCount).slice(0, 10);
    const top10ByTime = [...songRows].sort(compareSongsByListenTime).slice(0, 10);
    const top10Skipped = songRows
      .filter((row) => row.skipCount > 0)
      .sort(compareSongsBySkipCount)
      .slice(0, 10);

    const top3PerMonthRows = [];
    const sortedMonths = Array.from(monthlySongMap.keys()).sort((a, b) => a.localeCompare(b));
    for (const month of sortedMonths) {
      const monthSongs = monthlySongMap.get(month);
      if (!monthSongs) {
        continue;
      }

      const ranked = Array.from(monthSongs.entries())
        .sort((a, b) => compareMonthlySongRank(a, b))
        .slice(0, 3);

      ranked.forEach(([songKey, monthEntry], index) => {
        const [song, artist] = splitSongKey(songKey);
        top3PerMonthRows.push({
          month,
          monthLabel: formatMonthLabel(month),
          rank: index + 1,
          song,
          artist,
          playCount: monthEntry.playCount,
          minutesListened: monthEntry.totalMs / 60000,
        });
      });
    }

    const monthlyMinutesRows = Array.from(monthlyMs.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([month, totalMs]) => ({
        month,
        monthLabel: formatMonthLabel(month),
        totalMs,
        minutes: totalMs / 60000,
        hours: totalMs / 3600000,
      }));

    const yearlyMinutesRows = Array.from(yearlyMs.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([year, totalMs]) => ({
        year,
        totalMs,
        minutes: totalMs / 60000,
        hours: totalMs / 3600000,
      }));

    const peakMonth = pickPeak(monthlyMinutesRows, (row) => row.totalMs, (row) => row.month);
    const peakYear = pickPeak(yearlyMinutesRows, (row) => row.totalMs, (row) => String(row.year));

    const topArtistsByMinutes = Array.from(artistMap.values())
      .map((row) => ({
        artist: row.artist,
        playCount: row.playCount,
        skipCount: row.skipCount,
        totalMs: row.totalMs,
        totalMinutes: row.totalMs / 60000,
      }))
      .sort((a, b) => compareArtistsByTime(a, b))
      .slice(0, 10);

    const totalSongMs = songRows.reduce((sum, row) => sum + row.totalMs, 0);
    const dataRange = monthlyMinutesRows.length
      ? `${monthlyMinutesRows[0].monthLabel} - ${
          monthlyMinutesRows[monthlyMinutesRows.length - 1].monthLabel
        }`
      : "N/A";

    return {
      parsedInputMeta,
      totalRecordsRead: records.length,
      songRecordsUsed,
      ignoredNonSong,
      ignoredBadTimestamp,
      uniqueSongs: songRows.length,
      totalSongMs,
      totalMinutes: totalSongMs / 60000,
      totalHours: totalSongMs / 3600000,
      peakMonth,
      peakYear,
      dataRange,
      songRows,
      top10ByPlayCount,
      top10ByTime,
      top10Skipped,
      top3PerMonthRows,
      monthlyMinutesRows,
      yearlyMinutesRows,
      topArtistsByMinutes,
    };
  }

  function exportExcelReport(report) {
    const wb = XLSX.utils.book_new();
    const generatedUtc = new Date().toISOString().replace("T", " ").replace("Z", " UTC");

    addSheet(wb, "Summary", ["Metric", "Value"], [
      ["Generated At (UTC)", generatedUtc],
      ["JSON Files Parsed", report.parsedInputMeta.parsedJsonCount],
      ["Total Records Read", report.totalRecordsRead],
      ["Song Records Used", report.songRecordsUsed],
      ["Ignored Non-song Records", report.ignoredNonSong],
      ["Ignored Bad Timestamp Records", report.ignoredBadTimestamp],
      ["Unique Songs", report.uniqueSongs],
      ["Total Minutes Listened", round2(report.totalMinutes)],
      ["Total Hours Listened", round2(report.totalHours)],
      [
        "Month With Most Minutes",
        report.peakMonth
          ? `${report.peakMonth.month} (${round2(report.peakMonth.minutes)} min)`
          : "N/A",
      ],
      [
        "Year With Most Minutes",
        report.peakYear
          ? `${report.peakYear.year} (${round2(report.peakYear.minutes)} min)`
          : "N/A",
      ],
    ]);

    addSheet(
      wb,
      "Songs_Play_Counts",
      [
        "Song",
        "Artist",
        "Album",
        "Spotify Track URI",
        "Play Count",
        "Skip Count",
        "Skip Rate (0-1)",
        "Total Minutes Listened",
        "Total Hours Listened",
      ],
      report.songRows.map((row) => [
        row.song,
        row.artist,
        row.album,
        row.uri,
        row.playCount,
        row.skipCount,
        round4(row.skipRate),
        round2(row.totalMinutes),
        round2(row.totalHours),
      ])
    );

    addSheet(
      wb,
      "Top10_Play_Count",
      ["Rank", "Song", "Artist", "Play Count", "Total Minutes", "Skip Count"],
      report.top10ByPlayCount.map((row, index) => [
        index + 1,
        row.song,
        row.artist,
        row.playCount,
        round2(row.totalMinutes),
        row.skipCount,
      ])
    );

    addSheet(
      wb,
      "Top10_Listen_Time",
      ["Rank", "Song", "Artist", "Total Minutes", "Play Count", "Skip Count"],
      report.top10ByTime.map((row, index) => [
        index + 1,
        row.song,
        row.artist,
        round2(row.totalMinutes),
        row.playCount,
        row.skipCount,
      ])
    );

    addSheet(
      wb,
      "Top10_Skipped",
      ["Rank", "Song", "Artist", "Skip Count", "Play Count", "Skip Rate (0-1)", "Total Minutes"],
      report.top10Skipped.map((row, index) => [
        index + 1,
        row.song,
        row.artist,
        row.skipCount,
        row.playCount,
        round4(row.skipRate),
        round2(row.totalMinutes),
      ])
    );

    addSheet(
      wb,
      "Top3_Per_Month",
      ["Month", "Rank", "Song", "Artist", "Play Count", "Minutes Listened"],
      report.top3PerMonthRows.map((row) => [
        row.month,
        row.rank,
        row.song,
        row.artist,
        row.playCount,
        round2(row.minutesListened),
      ])
    );

    addSheet(
      wb,
      "Monthly_Minutes",
      ["Month", "Minutes Listened", "Hours Listened"],
      report.monthlyMinutesRows.map((row) => [row.month, round2(row.minutes), round2(row.hours)])
    );

    addSheet(
      wb,
      "Yearly_Minutes",
      ["Year", "Minutes Listened", "Hours Listened"],
      report.yearlyMinutesRows.map((row) => [row.year, round2(row.minutes), round2(row.hours)])
    );

    const filename = `spotify_report_full_${timestampForFilename()}.xlsx`;
    XLSX.writeFile(wb, filename, { compression: true });
    return filename;
  }

  async function exportSimplifiedPdf(report) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4", compress: true });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 38;
    const contentWidth = pageWidth - margin * 2;
    let cursorY = margin;

    const theme = {
      navy: [8, 32, 54],
      navySoft: [18, 56, 89],
      green: [43, 127, 102],
      textDark: [24, 42, 61],
      textLight: [234, 245, 255],
      line: [198, 217, 229],
    };

    drawHeader();
    drawSummaryCards();

    const monthlyChartUrl = await chartImage({
      type: "line",
      data: {
        labels: report.monthlyMinutesRows.map((row) => row.monthLabel),
        datasets: [
          {
            label: "Minutes listened",
            data: report.monthlyMinutesRows.map((row) => round2(row.minutes)),
            borderColor: "rgba(43,127,102,1)",
            backgroundColor: "rgba(43,127,102,0.2)",
            pointRadius: 2.2,
            borderWidth: 2,
            tension: 0.2,
            fill: true,
          },
        ],
      },
      options: chartOptions("Minutes"),
    });

    addSectionTitle("Monthly Listening Trend");
    reserveSpace(210);
    doc.addImage(monthlyChartUrl, "PNG", margin, cursorY, contentWidth, 190, undefined, "FAST");
    cursorY += 205;

    const artistsChartUrl = await chartImage({
      type: "bar",
      data: {
        labels: report.topArtistsByMinutes.slice(0, 8).map((row) => row.artist),
        datasets: [
          {
            label: "Minutes",
            data: report.topArtistsByMinutes.slice(0, 8).map((row) => round2(row.totalMinutes)),
            backgroundColor: [
              "rgba(62,157,127,0.85)",
              "rgba(56,145,117,0.85)",
              "rgba(49,132,106,0.85)",
              "rgba(42,118,95,0.85)",
              "rgba(36,106,85,0.85)",
              "rgba(31,93,75,0.85)",
              "rgba(26,83,66,0.85)",
              "rgba(21,73,58,0.85)",
            ],
            borderRadius: 8,
          },
        ],
      },
      options: chartOptions("Minutes", true),
    });

    addSectionTitle("Top Artists by Listening Time");
    reserveSpace(230);
    doc.addImage(artistsChartUrl, "PNG", margin, cursorY, contentWidth, 210, undefined, "FAST");
    cursorY += 226;

    reserveSpace(280, true);
    addSectionTitle("Top Songs Snapshot");
    doc.autoTable({
      startY: cursorY,
      margin: { left: margin, right: margin },
      head: [["Rank", "Song", "Artist", "Plays", "Minutes"]],
      body: report.top10ByPlayCount.slice(0, 10).map((row, index) => [
        index + 1,
        row.song,
        row.artist,
        row.playCount,
        round2(row.totalMinutes),
      ]),
      styles: { fontSize: 8.8, cellPadding: 4, textColor: theme.textDark },
      headStyles: { fillColor: theme.navySoft, textColor: theme.textLight },
      alternateRowStyles: { fillColor: [239, 247, 242] },
      tableLineColor: theme.line,
      tableLineWidth: 0.2,
      theme: "grid",
    });
    cursorY = doc.lastAutoTable.finalY + 14;

    reserveSpace(250, true);
    addSectionTitle("Most Skipped Songs");
    doc.autoTable({
      startY: cursorY,
      margin: { left: margin, right: margin },
      head: [["Rank", "Song", "Artist", "Skips", "Skip Rate"]],
      body: report.top10Skipped.slice(0, 10).map((row, index) => [
        index + 1,
        row.song,
        row.artist,
        row.skipCount,
        `${round2(row.skipRate * 100)}%`,
      ]),
      styles: { fontSize: 8.8, cellPadding: 4, textColor: theme.textDark },
      headStyles: { fillColor: theme.navySoft, textColor: theme.textLight },
      alternateRowStyles: { fillColor: [242, 248, 252] },
      tableLineColor: theme.line,
      tableLineWidth: 0.2,
      theme: "grid",
    });

    const filename = `spotify_report_snapshot_${timestampForFilename()}.pdf`;
    doc.save(filename);
    return filename;

    function drawHeader() {
      doc.setFillColor(...theme.navy);
      doc.roundedRect(margin, cursorY, contentWidth, 88, 14, 14, "F");

      doc.setFont("helvetica", "bold");
      doc.setTextColor(...theme.textLight);
      doc.setFontSize(21);
      doc.text("Spotify Listening Snapshot", margin + 18, cursorY + 34);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.text(`Generated: ${new Date().toISOString().replace("T", " ").slice(0, 19)} UTC`, margin + 18, cursorY + 56);
      doc.text(`Range: ${report.dataRange}`, margin + 18, cursorY + 72);

      cursorY += 108;
    }

    function drawSummaryCards() {
      const gap = 12;
      const cardW = (contentWidth - gap) / 2;
      const cardH = 72;

      const cards = [
        ["Total Minutes", formatNumber(round2(report.totalMinutes))],
        ["Unique Songs", formatNumber(report.uniqueSongs)],
        [
          "Highest Month",
          report.peakMonth
            ? `${report.peakMonth.monthLabel} (${formatNumber(round2(report.peakMonth.minutes))}m)`
            : "N/A",
        ],
        [
          "Highest Year",
          report.peakYear
            ? `${report.peakYear.year} (${formatNumber(round2(report.peakYear.minutes))}m)`
            : "N/A",
        ],
      ];

      cards.forEach((card, index) => {
        const row = Math.floor(index / 2);
        const col = index % 2;
        const x = margin + col * (cardW + gap);
        const y = cursorY + row * (cardH + gap);

        doc.setFillColor(238, 247, 255);
        doc.roundedRect(x, y, cardW, cardH, 10, 10, "F");
        doc.setDrawColor(189, 212, 230);
        doc.roundedRect(x, y, cardW, cardH, 10, 10, "S");

        doc.setFont("helvetica", "normal");
        doc.setTextColor(...theme.textDark);
        doc.setFontSize(10);
        doc.text(card[0], x + 12, y + 24);

        doc.setFont("helvetica", "bold");
        doc.setFontSize(14);
        doc.text(String(card[1]), x + 12, y + 47, { maxWidth: cardW - 20 });
      });

      cursorY += cardH * 2 + gap + 18;
    }

    function addSectionTitle(title) {
      reserveSpace(26, true);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.setTextColor(...theme.navy);
      doc.text(title, margin, cursorY);
      cursorY += 10;
    }

    function reserveSpace(height, allowNewPage) {
      if (cursorY + height <= pageHeight - margin) {
        return;
      }
      if (!allowNewPage) {
        return;
      }
      doc.addPage();
      cursorY = margin;
    }
  }

  function chartOptions(yTitle, horizontal = false) {
    return {
      responsive: false,
      animation: false,
      indexAxis: horizontal ? "y" : "x",
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
      },
      scales: {
        x: {
          ticks: { color: "#2b3d4f", font: { size: 10 } },
          grid: { color: "rgba(82,117,143,0.2)" },
        },
        y: {
          ticks: { color: "#2b3d4f", font: { size: 10 } },
          title: {
            display: Boolean(yTitle),
            text: yTitle,
            color: "#2b3d4f",
            font: { size: 10 },
          },
          grid: { color: "rgba(82,117,143,0.2)" },
          beginAtZero: true,
        },
      },
    };
  }

  async function chartImage(config) {
    const canvas = document.createElement("canvas");
    canvas.width = 1200;
    canvas.height = 560;

    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Could not create chart canvas context.");
    }

    const chart = new Chart(context, config);
    await new Promise((resolve) => setTimeout(resolve, 40));
    const url = canvas.toDataURL("image/png");
    chart.destroy();
    return url;
  }

  function addSheet(workbook, title, headers, rows) {
    const data = [headers, ...rows];
    const sheet = XLSX.utils.aoa_to_sheet(data);
    sheet["!cols"] = columnWidths(data);

    if (data.length > 0 && headers.length > 0) {
      sheet["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: data.length - 1, c: headers.length - 1 },
        }),
      };
    }

    XLSX.utils.book_append_sheet(workbook, sheet, uniqueSheetName(workbook, title));
  }

  function uniqueSheetName(workbook, desiredName) {
    const used = new Set(workbook.SheetNames);
    let base = desiredName.replace(/[\\/*?:[\]]/g, "_").trim();
    if (!base) {
      base = "Sheet";
    }
    base = base.slice(0, 31);

    if (!used.has(base)) {
      return base;
    }

    let index = 2;
    while (true) {
      const suffix = `_${index}`;
      const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
      if (!used.has(candidate)) {
        return candidate;
      }
      index += 1;
    }
  }

  function columnWidths(table) {
    if (!table.length) {
      return [];
    }
    const widthCount = table[0].length;
    const widths = new Array(widthCount).fill(10);

    for (const row of table) {
      for (let col = 0; col < widthCount; col += 1) {
        const value = row[col];
        const length = String(value ?? "").length;
        if (length > widths[col]) {
          widths[col] = Math.min(length + 2, 70);
        }
      }
    }

    return widths.map((wch) => ({ wch }));
  }

  function isSongRecord(record) {
    const trackName = cleanString(record.master_metadata_track_name);
    const artistName = cleanString(record.master_metadata_album_artist_name);
    const episodeName = cleanString(record.episode_name);
    const episodeUri = cleanString(record.spotify_episode_uri);

    if (!trackName || !artistName) {
      return false;
    }
    if (episodeName || episodeUri) {
      return false;
    }
    return true;
  }

  function parseTimestamp(value) {
    if (typeof value !== "string" || !value.trim()) {
      return null;
    }

    let date = new Date(value);
    if (!Number.isNaN(date.getTime())) {
      return date;
    }

    const maybeIso = value.includes("T") ? value : value.replace(" ", "T");
    const withZone = maybeIso.endsWith("Z") ? maybeIso : `${maybeIso}Z`;
    date = new Date(withZone);
    if (Number.isNaN(date.getTime())) {
      return null;
    }
    return date;
  }

  function timestampToMonthKey(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  function cleanString(value) {
    if (typeof value !== "string") {
      return "";
    }
    return value.trim();
  }

  function sanitizeMs(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) {
      return 0;
    }
    if (num < 0) {
      return 0;
    }
    return Math.round(num);
  }

  function incrementCounter(counter, value) {
    if (!value) {
      return;
    }
    counter.set(value, (counter.get(value) || 0) + 1);
  }

  function mostCommonCounterKey(counter) {
    let bestKey = "";
    let bestCount = -1;
    for (const [key, count] of counter.entries()) {
      if (count > bestCount) {
        bestKey = key;
        bestCount = count;
      }
    }
    return bestKey;
  }

  function makeSongKey(song, artist) {
    return `${song}${KEY_SEP}${artist}`;
  }

  function splitSongKey(songKey) {
    const splitIndex = songKey.indexOf(KEY_SEP);
    if (splitIndex < 0) {
      return [songKey, ""];
    }
    return [songKey.slice(0, splitIndex), songKey.slice(splitIndex + 1)];
  }

  function compareSongsByPlayCount(a, b) {
    return (
      (b.playCount - a.playCount) ||
      (b.totalMs - a.totalMs) ||
      (b.skipCount - a.skipCount) ||
      a.artist.localeCompare(b.artist) ||
      a.song.localeCompare(b.song)
    );
  }

  function compareSongsByListenTime(a, b) {
    return (
      (b.totalMs - a.totalMs) ||
      (b.playCount - a.playCount) ||
      a.artist.localeCompare(b.artist) ||
      a.song.localeCompare(b.song)
    );
  }

  function compareSongsBySkipCount(a, b) {
    return (
      (b.skipCount - a.skipCount) ||
      (b.playCount - a.playCount) ||
      a.artist.localeCompare(b.artist) ||
      a.song.localeCompare(b.song)
    );
  }

  function compareMonthlySongRank(a, b) {
    const [songA, statA] = a;
    const [songB, statB] = b;
    const [titleA, artistA] = splitSongKey(songA);
    const [titleB, artistB] = splitSongKey(songB);

    return (
      (statB.playCount - statA.playCount) ||
      (statB.totalMs - statA.totalMs) ||
      artistA.localeCompare(artistB) ||
      titleA.localeCompare(titleB)
    );
  }

  function compareArtistsByTime(a, b) {
    return (
      (b.totalMs - a.totalMs) ||
      (b.playCount - a.playCount) ||
      a.artist.localeCompare(b.artist)
    );
  }

  function pickPeak(rows, valueAccessor, tieBreakerAccessor) {
    if (!rows.length) {
      return null;
    }
    let best = rows[0];

    for (let index = 1; index < rows.length; index += 1) {
      const current = rows[index];
      const currentValue = valueAccessor(current);
      const bestValue = valueAccessor(best);

      if (
        currentValue > bestValue ||
        (currentValue === bestValue &&
          String(tieBreakerAccessor(current)) > String(tieBreakerAccessor(best)))
      ) {
        best = current;
      }
    }

    if ("month" in best) {
      return {
        month: best.month,
        monthLabel: best.monthLabel,
        totalMs: best.totalMs,
        minutes: best.minutes,
      };
    }

    return {
      year: best.year,
      totalMs: best.totalMs,
      minutes: best.minutes,
    };
  }

  function isSupportedInputFile(fileName) {
    const lower = fileName.toLowerCase();
    return lower.endsWith(".zip") || lower.endsWith(".json");
  }

  function getSelectedMode() {
    const selected = document.querySelector("input[name='outputMode']:checked");
    return selected?.value === "pdf" ? "pdf" : "excel";
  }

  function fileUniqueKey(file) {
    return `${file.name}::${file.size}::${file.lastModified}`;
  }

  function formatFileSize(bytes) {
    const units = ["B", "KB", "MB", "GB"];
    let size = bytes;
    let unitIndex = 0;
    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex += 1;
    }
    return `${size.toFixed(size < 10 && unitIndex > 0 ? 1 : 0)} ${units[unitIndex]}`;
  }

  function formatMonthLabel(monthKey) {
    const parts = monthKey.split("-");
    if (parts.length !== 2) {
      return monthKey;
    }
    const year = Number(parts[0]);
    const month = Number(parts[1]);
    if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) {
      return monthKey;
    }
    return `${MONTH_NAMES[month - 1]} ${year}`;
  }

  function round2(value) {
    return Math.round((value + Number.EPSILON) * 100) / 100;
  }

  function round4(value) {
    return Math.round((value + Number.EPSILON) * 10000) / 10000;
  }

  function formatNumber(value) {
    return Number(value).toLocaleString();
  }

  function timestampForFilename() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const min = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    return `${yyyy}${mm}${dd}_${hh}${min}${ss}`;
  }

  function addStatus(message, level = "info") {
    const existingWaiting = dom.statusLog.querySelector("li");
    if (existingWaiting && existingWaiting.textContent === "Waiting for files.") {
      dom.statusLog.innerHTML = "";
    }

    const item = document.createElement("li");
    item.textContent = message;
    if (level === "ok") {
      item.classList.add("ok");
    } else if (level === "warn") {
      item.classList.add("warn");
    } else if (level === "error") {
      item.classList.add("error");
    }
    dom.statusLog.appendChild(item);
    item.scrollIntoView({ block: "nearest" });
  }

  function resetStatusLog() {
    dom.statusLog.innerHTML = "";
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function errorMessage(error) {
    if (error instanceof Error) {
      return error.message;
    }
    return String(error);
  }
})();
