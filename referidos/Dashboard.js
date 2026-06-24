var SHEET_ID = '1Sdy_XcAK8MBsx1YrDZosGaNcG4wg7LKlddAVcdfIplY';
var SHEET_NAME = 'referrals';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dashboard Médicos Referentes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var rows = data.slice(1);
  var patients = [];

  rows.forEach(function(row) {
    var title = row[1] || '';
    var stage = row[2] || '';
    var status = row[3] || '';
    var etiqueta = row[4] || '';
    var doctor = row[5] || '';
    var addTime = row[6] || '';
    var referralDate = row[7] || '';

    if (!title && !doctor) return;

    patients.push({
      id: formatPatientId(title),
      stage: stage.toString().trim(),
      status: status.toString().trim().toLowerCase(),
      campaign: etiqueta.toString().trim() || 'Sin etiqueta',
      doctor: doctor.toString().trim(),
      addTime: formatDate(addTime),
      referralDate: formatDate(referralDate)
    });
  });

  return JSON.stringify(buildSummary(patients));
}

function formatPatientId(title) {
  var s = title.toString().trim();
  var parts = s.split('-');
  if (parts.length > 1) return parts[parts.length - 1].substring(0, 6);
  return s.slice(-6);
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  }
  // Try to parse string dates (may include timestamp)
  var s = val.toString().trim();
  if (!s) return '';
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    }
  } catch(e) {}
  // Last resort: take only the date part if there's a space
  return s.split(' ')[0];
}

function buildSummary(patients) {
  var campaigns = {};
  var doctors = {};
  var monthlyRef = {};
  var monthlyDeriv = {};

  // Sort patients newest first by addTime
  patients.sort(function(a, b) {
    if (!a.addTime) return 1;
    if (!b.addTime) return -1;
    return b.addTime > a.addTime ? 1 : -1;
  });

  patients.forEach(function(p) {
    if (!campaigns[p.campaign]) {
      campaigns[p.campaign] = {
        name: p.campaign, doctors: {}, total: 0, won: 0, lost: 0, open: 0,
        stages: {}, lostByStage: {}, monthly: {}, latestDate: ''
      };
    }
    var c = campaigns[p.campaign];
    c.total++;
    if (p.status === 'won') c.won++;
    else if (p.status === 'lost') {
      c.lost++;
      c.lostByStage[p.stage] = (c.lostByStage[p.stage] || 0) + 1;
    } else {
      c.open++;
    }
    c.doctors[p.doctor] = true;
    c.stages[p.stage] = (c.stages[p.stage] || 0) + 1;

    // Track latest patient date per campaign
    if (p.addTime && p.addTime > c.latestDate) c.latestDate = p.addTime;

    // Monthly progression (open + won only)
    if (p.addTime && p.status !== 'lost') {
      var mo = p.addTime.substring(0, 7);
      if (!c.monthly[mo]) c.monthly[mo] = { open: 0, won: 0 };
      if (p.status === 'won') c.monthly[mo].won++;
      else c.monthly[mo].open++;
    }

    if (!doctors[p.doctor]) {
      doctors[p.doctor] = { name: p.doctor, campaigns: {}, total: 0 };
    }
    doctors[p.doctor].total++;
    doctors[p.doctor].campaigns[p.campaign] = true;

    if (p.addTime) {
      var mRef = p.addTime.substring(0, 7);
      monthlyRef[mRef] = (monthlyRef[mRef] || 0) + 1;
    }
    if (p.referralDate) {
      var mDeriv = p.referralDate.substring(0, 7);
      monthlyDeriv[mDeriv] = (monthlyDeriv[mDeriv] || 0) + 1;
    }
  });

  // Build campaign list sorted by latest patient date (newest first)
  var campaignList = Object.values(campaigns).map(function(c) {
    var months = Object.keys(c.monthly).sort();
    return {
      name: c.name,
      doctorCount: Object.keys(c.doctors).length,
      total: c.total, won: c.won, lost: c.lost, open: c.open,
      stages: c.stages,
      lostByStage: c.lostByStage,
      latestDate: c.latestDate,
      monthly: {
        labels: months,
        open: months.map(function(m) { return c.monthly[m].open; }),
        won: months.map(function(m) { return c.monthly[m].won; })
      }
    };
  }).sort(function(a, b) {
    if (!a.latestDate) return 1;
    if (!b.latestDate) return -1;
    return b.latestDate > a.latestDate ? 1 : -1;
  });

  return {
    patients: patients,
    campaigns: campaignList,
    doctors: Object.values(doctors).sort(function(a, b) { return b.total - a.total; }),
    monthlyRef: sortMonthly(monthlyRef),
    monthlyDeriv: sortMonthly(monthlyDeriv),
    totals: {
      patients: patients.length,
      won: patients.filter(function(p) { return p.status === 'won'; }).length,
      lost: patients.filter(function(p) { return p.status === 'lost'; }).length,
      open: patients.filter(function(p) { return p.status === 'open'; }).length,
      doctors: Object.keys(doctors).length,
      campaigns: campaignList.length
    },
    lastUpdated: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm')
  };
}

function sortMonthly(obj) {
  var keys = Object.keys(obj).sort();
  return { labels: keys, values: keys.map(function(k) { return obj[k]; }) };
}
