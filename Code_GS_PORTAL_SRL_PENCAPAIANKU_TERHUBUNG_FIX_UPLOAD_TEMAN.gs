/**
 * BACKEND PORTAL-SRL (Portofolio Digital Sekolah Ramah Lingkungan)
 * Terhubung ke 3 Tab Google Sheets: USERS, STUDENT_SUBMISSIONS, TEACHER_EVALUATIONS
 */

// Helper untuk mendapatkan objek Spreadsheet secara aman dan efisien
function getSS() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    const id = typeof SPREADSHEET_ID !== 'undefined' ? SPREADSHEET_ID : '';
    return SpreadsheetApp.openById(id);
  }
}

// Helper untuk memastikan Tab & Header Sheet tersedia (Self-healing)
function getOrCreateSheet(sheetName) {
  const ss = getSS();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === 'USERS') {
      sheet.appendRow(['ID_NISN', 'NAMA', 'ROLE', 'KELAS', 'ABSEN', 'PASSWORD']);
    } else if (sheetName === 'STUDENT_SUBMISSIONS') {
      sheet.appendRow([
        'submission_id', 'timestamp', 'nisn', 'act_id',
        'jawaban_teks_1', 'jawaban_teks_2', 'link_karya_padlet',
        'peer_assessment', 'self_assessment', 'refleksi_murid', 'status_misi'
      ]);
    } else if (sheetName === 'TEACHER_EVALUATIONS') {
      sheet.appendRow([
        'EVAL_ID', 'TIMESTAMP', 'NISN', 'ACT_ID', 
        'RUBRIK_PENILAIAN', 'OBSERVASI_GURU', 'APK_PENALARAN', 
        'APK_KREATIVITAS', 'APK_KOLABORASI', 'FEEDBACK_GURU', 'TINDAK_LANJUT'
      ]);
    } else if (sheetName === 'Pengaturan') {
      sheet.appendRow(['KEY', 'VALUE', 'UPDATED_AT']);
    } else if (sheetName === 'Pengaturan_Aktivitas') {
      sheet.appendRow(['AKTIVITAS_ID', 'STATUS', 'UPDATED_AT']);
    } else if (sheetName === 'TEACHER_FEEDBACK') {
      sheet.appendRow(['FEEDBACK_ID','TIMESTAMP','STUDENT_NISN','STUDENT_NAME','STUDENT_CLASS','ACT_ID','TYPE','MESSAGE','TEACHER_NIP','TEACHER_NAME']);
    }
  }
  return sheet;
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let responseData = { status: 'error', message: 'Aksi tidak dikenali' };

  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        // Payload tetap objek kosong jika parsing gagal
      }
    }

    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || payload.action;

    switch (action) {
      case 'login':
        responseData = handleLogin(
          payload.user_id || payload.username || params.user_id || params.username,
          payload.password || params.password
        );
        break;

      case 'loginGuru':
        responseData = handleLoginGuru(
          payload.username || payload.user_id || params.username || params.user_id,
          payload.password || params.password
        );
        break;

      case 'registerGuru':
        responseData = registerGuru(payload.nip || params.nip, payload.password !== undefined ? payload.password : params.password);
        break;

      case 'verifyStudentLogin':
        responseData = verifyStudentLogin(
          payload.nama || params.nama,
          payload.kelas || params.kelas,
          payload.absen || params.absen
        );
        break;

      case 'getStudentData':
        responseData = getStudentData(payload.nisn || params.nisn);
        break;

      case 'getStudentClassmates':
        responseData = getStudentClassmates(payload.nisn || params.nisn);
        break;

      case 'getUserStateData':
      case 'fetchInitialData':
        responseData = getUserStateData(payload.id_nisn || payload.nisn || params.id_nisn || params.nisn);
        break;

      case 'saveStudentSubmission':
        responseData = saveStudentSubmission(payload);
        break;

      case 'getTeacherNotifications':
        responseData = getTeacherNotifications(payload);
        break;

      case 'markTeacherNotificationRead':
        responseData = markTeacherNotificationRead(payload);
        break;

      case 'saveTeacherFeedback':
        responseData = saveTeacherFeedback(payload);
        break;

      case 'getStudentFeedback':
        responseData = getStudentFeedback(payload.nisn || params.nisn);
        break;

      case 'uploadFileToDrive':
        responseData = uploadFileToDrive(payload);
        break;

      case 'addStudent':
        responseData = addStudent(payload.studentData || payload);
        break;

      case 'getTeacherData':
        responseData = getTeacherData();
        break;

      case 'getInitialAppData':
        responseData = getInitialAppData(!!payload.forceRefresh);
        break;

      case 'getPengaturan':
      case 'getTeacherSettings':
        responseData = getPengaturan();
        break;

      case 'getPublishedPengantar':
        responseData = getPublishedPengantar();
        break;

      case 'savePengaturan':
      case 'saveTeacherSettings':
        responseData = savePengaturan(payload.teacherSettings || payload.data || payload);
        break;

      case 'saveRestrictionStatus':
        responseData = saveRestrictionStatus(payload.aktivitasId || payload.act_id, payload.status);
        break;

      case 'getRestrictionStatus':
        responseData = getRestrictionStatus();
        break;

      case 'getDashboardStats':
        responseData = getDashboardStats();
        break;

      case 'saveTeacherEvaluation':
        responseData = saveTeacherEvaluation(payload);
        break;

      case 'saveTeacherActivityAssessment':
        responseData = saveTeacherActivityAssessment(payload);
        break;

      case 'getStudentAPKProfile':
        responseData = getStudentAPKProfile(payload.nisn || params.nisn);
        break;

      case 'getTeacherEvaluation':
        responseData = getTeacherEvaluation(payload.nisn || params.nisn, payload.act_id || params.act_id);
        break;

      default:
        responseData = { status: 'error', message: 'Aksi parameter tidak valid atau tidak ditemukan: ' + action };
    }
  } catch (err) {
    responseData = { status: 'error', message: 'Error Server GAS: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Cache data awal aplikasi agar pembukaan portal lebih cepat.
 * Cache dibersihkan setiap kali data profil, akses, murid, bukti, atau evaluasi berubah.
 */
const PORTAL_INITIAL_APP_CACHE_KEY = 'PORTAL_SRL_INITIAL_APP_DATA_V1';
const PORTAL_INITIAL_APP_CACHE_TTL_SECONDS = 30;

function invalidateInitialAppDataCache_() {
  try { CacheService.getScriptCache().remove(PORTAL_INITIAL_APP_CACHE_KEY); } catch (e) {}
}

/**
 * Mengambil profil/pengaturan, statistik beranda, dan status Restrict Akses
 * dalam satu panggilan API frontend. Data Sheets dibaca melalui fungsi batch.
 */
function getInitialAppData(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    try {
      const cached = cache.get(PORTAL_INITIAL_APP_CACHE_KEY);
      if (cached) {
        const parsed = JSON.parse(cached);
        parsed.fromCache = true;
        return parsed;
      }
    } catch (e) {}
  }

  try {
    // Tiap helper mengambil data per sheet dengan getDataRange().getValues(), bukan per sel.
    const settingsResult = getPengaturan();
    const statsResult = getDashboardStats();
    const restrictionsResult = getRestrictionStatus();

    if (settingsResult.status !== 'success') throw new Error(settingsResult.message || 'Profil tidak dapat dibaca.');
    if (statsResult.status !== 'success') throw new Error(statsResult.message || 'Statistik tidak dapat dihitung.');
    if (restrictionsResult.status !== 'success') throw new Error(restrictionsResult.message || 'Status aktivitas tidak dapat dibaca.');

    const result = {
      status: 'success',
      settings: settingsResult.settings || {},
      profile: {
        teacherDisplayName: (settingsResult.settings || {}).teacherDisplayName || '',
        teacherUsername: (settingsResult.settings || {}).teacherUsername || '',
        teacherRole: (settingsResult.settings || {}).teacherRole || (settingsResult.settings || {}).teacherClass || '',
        schoolName: (settingsResult.settings || {}).schoolName || ''
      },
      stats: statsResult.stats || {},
      restrictions: restrictionsResult.restrictions || {},
      fetchedAt: new Date().toISOString(),
      fromCache: false
    };

    try { cache.put(PORTAL_INITIAL_APP_CACHE_KEY, JSON.stringify(result), PORTAL_INITIAL_APP_CACHE_TTL_SECONDS); } catch (e) {}
    return result;
  } catch (err) {
    return { status: 'error', message: 'Gagal memuat data awal PORTAL-SRL: ' + err.message };
  }
}

/**
 * Mengirim pengantar Topik 2 yang dipublikasikan dari portal Guru ke portal Murid.
 * Pengaturan disimpan melalui savePengaturan() dalam bentuk JSON pada kolom VALUE.
 */
function getPublishedPengantar() {
  try {
    const result = getPengaturan();
    if (!result || result.status !== 'success') {
      return { status: 'error', message: result && result.message ? result.message : 'Pengaturan guru belum dapat dibaca.' };
    }
    const settings = result.settings || {};
    const content = settings.pengantarContent;
    const topic = settings.selectedTopik;
    return {
      status: 'success',
      teacherSettings: {
        pengantarContent: content == null || String(content).trim() === ''
          ? 'Pada Topik 2 ini, kamu diajak menjadi anak yang peduli dan bertanggung jawab terhadap kebersihan lingkungan. Setiap hari, kita menghasilkan sampah dari berbagai aktivitas, seperti makan, belajar, dan bermain. Jika sampah tidak dikelola dengan baik, lingkungan dapat menjadi kotor, tidak sehat, serta mengganggu kenyamanan dan kesehatan semua orang.\n\nMelalui kegiatan pada topik ini, kamu akan belajar mengenal jenis-jenis sampah, seperti sampah organik, anorganik, B3, dan sampah elektronik, serta memahami dampaknya bagi lingkungan. Kamu juga diajak mengelola sampah dengan cara sederhana, bekerja sama dalam kelompok, berpikir kritis terhadap kebiasaan sehari-hari, dan melakukan aksi nyata untuk menjaga lingkungan. Dengan demikian, kamu diharapkan mampu membangun sikap peduli lingkungan secara berkelanjutan dalam kehidupan sehari-hari.'
          : String(content),
        selectedTopik: topic == null || String(topic).trim() === ''
          ? 'Topik 2: Aku Bisa Menjaga Lingkungan'
          : String(topic)
      }
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal mengambil pengantar SRL: ' + err.message };
  }
}

/** Membaca pengaturan profil guru dan sekolah dari sheet Pengaturan. */
function getPengaturan() {
  try {
    const sheet = getOrCreateSheet('Pengaturan');
    const values = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      if (!key) continue;
      try { settings[key] = JSON.parse(String(values[i][1] ?? 'null')); }
      catch (e) { settings[key] = values[i][1]; }
    }
    return { status: 'success', settings: settings, updatedAt: values.length > 1 ? values[values.length - 1][2] : null };
  } catch (err) {
    return { status: 'error', message: 'Gagal membaca pengaturan: ' + err.message };
  }
}

/** Menyimpan pengaturan profil guru/sekolah sebagai key-value agar mudah diperluas. */
function savePengaturan(data) {
  if (!data || typeof data !== 'object') return { status: 'error', message: 'Data pengaturan tidak valid.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateSheet('Pengaturan');
    const values = sheet.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      if (key) rowByKey[key] = i + 1;
    }
    // Hanya simpan pengaturan portal yang diharapkan; password/kredensial tidak disimpan di sini.
    const allowed = [
      'schoolName','npsn','schoolAddress','principal','schoolYear','schoolDescription',
      'teacherDisplayName','teacherUsername','teacherRole','teacherClass','selectedTopik',
      'headerImage','logoUrl','themeColor','pengantarContent','stimulasiPrompt',
      'stimulasiImage','stimulasiVideo','unlockedActivities'
    ];
    const now = new Date();
    allowed.forEach(function(key) {
      if (!Object.prototype.hasOwnProperty.call(data, key)) return;
      const encoded = JSON.stringify(data[key] === undefined ? '' : data[key]);
      const row = rowByKey[key];
      if (row) sheet.getRange(row, 2, 1, 2).setValues([[encoded, now]]);
      else sheet.appendRow([key, encoded, now]);
    });
    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();
    return { status: 'success', message: 'Pengaturan berhasil disimpan.', settings: getPengaturan().settings };
  } catch (err) {
    return { status: 'error', message: 'Gagal menyimpan pengaturan: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Pendaftaran akun guru. USERS memiliki urutan kolom:
// ID_NISN, NAMA, ROLE, KELAS, ABSEN, PASSWORD.
function registerGuru(nip, password) {
  const cleanNip = String(nip || '').trim();
  // Password sengaja tidak di-trim agar nilai yang disimpan sama persis dengan input.
  const rawPassword = password === undefined || password === null ? '' : String(password);

  if (!cleanNip || rawPassword.length === 0) {
    return { status: 'error', message: 'NIP dan Password wajib diisi.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateSheet('USERS');
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      const duplicate = ids.some(row => String(row[0] || '').trim().toLowerCase() === cleanNip.toLowerCase());
      if (duplicate) {
        return { status: 'error', message: 'NIP sudah terdaftar.' };
      }
    }

    // Kolom tetap mengikuti skema USERS yang dipakai fitur login dan dashboard.
    // NAMA, KELAS, dan ABSEN dikosongkan karena form saat ini hanya meminta NIP/password.
    sheet.appendRow([cleanNip, '', 'GURU', '', '', rawPassword]);
    return { status: 'success', message: 'Akun Guru berhasil dibuat.' };
  } catch (err) {
    return { status: 'error', message: 'Pendaftaran gagal diproses: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

// Helper penyamaan format Kelas (misal: "Kelas IV SD" == "Kelas 4 SD")
function normalizeKelas(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .replace(/kelas/g, '')
    .replace(/sd/g, '')
    .replace(/iv/g, '4')
    .replace(/v/g, '5')
    .replace(/vi/g, '6')
    .replace(/\s+/g, '');
}

// Handler Otentikasi Login Guru
function handleLoginGuru(username, password) {
  const sheet = getOrCreateSheet('USERS');
  const data = sheet.getDataRange().getValues();
  const cleanUsername = String(username || '').trim().toLowerCase();
  const cleanPassword = String(password || '').trim();

  if (!cleanUsername || !cleanPassword) {
    return { status: 'error', message: 'NIP/NUPTK dan Password wajib diisi!' };
  }

  for (let i = 1; i < data.length; i++) {
    const rowUserId = String(data[i][0] || '').trim().toLowerCase();
    const rowNama = String(data[i][1] || '').trim();
    const rowRole = String(data[i][2] || '').trim().toLowerCase();
    const rowPassword = String(data[i][5] !== undefined ? data[i][5] : '').trim();

    if (rowRole === 'guru' || rowRole === 'teacher') {
      if (rowUserId === cleanUsername || rowNama.toLowerCase() === cleanUsername) {
        if (rowPassword === cleanPassword) {
          const teacherToken = createTeacherSession(data[i][0]);
          return {
            status: 'success',
            teacherToken: teacherToken,
            nama: rowNama,
            namaGuru: rowNama,
            data: {
              user_id: data[i][0],
              nama: rowNama,
              role: rowRole,
              kelas: data[i][3],
              absen: data[i][4]
            },
            message: 'Login Guru Berhasil!'
          };
        } else {
          return { status: 'error', message: 'Password guru tidak sesuai!' };
        }
      }
    }
  }

  return { status: 'error', message: 'NIP/NUPTK atau Username Guru tidak ditemukan di tab USERS!' };
}

// Handler Login Umum
function handleLogin(userId, password) {
  const sheet = getOrCreateSheet('USERS');
  const data = sheet.getDataRange().getValues();
  const cleanUserId = String(userId || '').trim().toLowerCase();
  const cleanPassword = String(password || '').trim();

  for (let i = 1; i < data.length; i++) {
    const rowUserId = String(data[i][0] || '').trim().toLowerCase();
    const rowNama = String(data[i][1] || '').trim();
    const rowRole = String(data[i][2] || '').trim().toLowerCase();
    const rowKelas = data[i][3];
    const rowAbsen = data[i][4];
    const rowPassword = String(data[i][5] !== undefined ? data[i][5] : '').trim();

    if (rowUserId === cleanUserId || rowNama.toLowerCase() === cleanUserId) {
      if ((rowRole === 'guru' || rowRole === 'teacher') && rowPassword !== cleanPassword) {
        return { status: 'error', message: 'Password guru tidak sesuai!' };
      }
      return {
        status: 'success',
        nama: rowNama,
        data: {
          user_id: data[i][0],
          nama: rowNama,
          role: rowRole,
          kelas: rowKelas,
          absen: rowAbsen
        }
      };
    }
  }
  return { status: 'error', message: 'ID Pengguna / NISN tidak ditemukan' };
}

// Verifikasi Login Murid (Nama, Kelas, Absen)
function verifyStudentLogin(nama, kelas, absen) {
  const sheet = getOrCreateSheet('USERS');
  const data = sheet.getDataRange().getValues();

  if (!nama) {
    return { status: 'error', message: 'Nama lengkap wajib diisi!' };
  }

  const cleanNama = String(nama).trim().toLowerCase();
  const cleanKelasNorm = normalizeKelas(kelas);
  const cleanAbsen = String(absen || '').trim();

  let nameMatch = null;

  for (let i = 1; i < data.length; i++) {
    const rowUserId = String(data[i][0]).trim();
    const rowNama = String(data[i][1]).trim();
    const rowRole = String(data[i][2] || '').trim().toLowerCase();
    const rowKelas = String(data[i][3]).trim();
    const rowAbsen = String(data[i][4] || '').trim();

    if ((rowRole === 'murid' || rowRole === 'siswa') && rowNama.toLowerCase() === cleanNama) {
      nameMatch = {
        nisn: rowUserId,
        nama: rowNama,
        kelas: rowKelas,
        absen: rowAbsen
      };

      const rowKelasNorm = normalizeKelas(rowKelas);

      if (rowKelasNorm === cleanKelasNorm && (!cleanAbsen || rowAbsen === cleanAbsen)) {
        return {
          status: 'success',
          student: nameMatch
        };
      }
    }
  }

  if (!nameMatch) {
    return { status: 'error', message: `Nama "${nama}" tidak ditemukan dalam database USERS.` };
  } else if (normalizeKelas(nameMatch.kelas) !== cleanKelasNorm) {
    return { status: 'error', message: `Kelas tidak sesuai. Murid atas nama ${nameMatch.nama} terdaftar di ${nameMatch.kelas}.` };
  } else if (cleanAbsen && nameMatch.absen !== cleanAbsen) {
    return { status: 'error', message: `Nomor absen salah. Absen yang terdaftar untuk ${nameMatch.nama} adalah ${nameMatch.absen}.` };
  }

  return { status: 'success', student: nameMatch };
}

// Mengambil teman sekelas yang benar-benar terdaftar pada tab USERS.
// Data contoh pada frontend tidak digunakan untuk penilaian teman sebaya.
function getStudentClassmates(nisn) {
  const cleanNisn = String(nisn || '').trim();
  if (!cleanNisn) return { status: 'error', message: 'NISN murid wajib diisi.' };

  const usersSheet = getOrCreateSheet('USERS');
  const rows = usersSheet.getDataRange().getValues();
  let currentStudent = null;
  const students = [];

  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '').trim();
    const name = String(rows[i][1] || '').trim();
    const role = String(rows[i][2] || '').trim().toLowerCase();
    const kelas = String(rows[i][3] || '').trim();
    const absen = String(rows[i][4] || '').trim();
    if (!id || !name || !['murid', 'siswa'].includes(role)) continue;
    const student = { nisn: id, nama: name, kelas: kelas, absen: absen };
    students.push(student);
    if (id === cleanNisn) currentStudent = student;
  }

  if (!currentStudent) {
    return { status: 'error', message: 'Data murid yang sedang login tidak ditemukan pada USERS.' };
  }
  const currentClass = normalizeKelas(currentStudent.kelas);
  if (!currentClass) return { status: 'success', classmates: [] };

  const classmates = students.filter(s =>
    s.nisn !== cleanNisn && normalizeKelas(s.kelas) === currentClass
  ).map(s => ({ nisn: s.nisn, nama: s.nama, kelas: s.kelas, absen: s.absen }));

  return { status: 'success', classmates: classmates };
}

// Ambil Data Progres & Jawaban Murid dari Sheet
function getStudentData(nisn) {
  const sheetSub = getOrCreateSheet('STUDENT_SUBMISSIONS');
  const dataSub = sheetSub.getDataRange().getValues();
  const answers = {};
  const progress = { 
    pengantar: true, 
    stimulasi: false, 
    act1: false, act2: false, act3: false, act4: false, act5: false,
    refleksi_akhir: false
  };
  
  const cleanNisn = String(nisn || '').trim();

  for (let i = 1; i < dataSub.length; i++) {
    if (String(dataSub[i][2] || '').trim() === cleanNisn) {
      const actId = String(dataSub[i][3] || '').toLowerCase();
      progress[actId] = (String(dataSub[i][10] || '').toLowerCase() === 'selesai');

      answers[`${actId}_jawaban_1`] = dataSub[i][4] || '';
      answers[`${actId}_jawaban_2`] = dataSub[i][5] || '';
      answers[`${actId}_padlet`] = dataSub[i][6] || '';
      answers[`${actId}_peer`] = dataSub[i][7] || '';
      answers[`${actId}_self`] = dataSub[i][8] || '';
      answers[`${actId}_refleksi`] = dataSub[i][9] || '';

      // Pemetaan Kunci Spesifik Halaman Frontend
      if (actId === 'stimulasi') {
        answers['stimulasi_respon'] = dataSub[i][4] || '';
      }
      if (actId === 'act1') answers['act1_identifikasi'] = dataSub[i][4] || '';
      if (actId === 'act2') answers['act2_dampak'] = dataSub[i][4] || '';
      if (actId === 'act3') {
        answers['act3_organik'] = dataSub[i][4] || '';
        answers['act3_anorganik'] = dataSub[i][5] || '';
      }
      if (actId === 'act4') {
        answers['act4_rancangan'] = dataSub[i][4] || '';
        answers['act4_link_video'] = dataSub[i][5] || dataSub[i][6] || '';
      }
      if (actId === 'act5') answers['act5_kebiasaan'] = dataSub[i][4] || '';
      if (actId === 'refleksi_akhir') {
        answers['refleksi_akhir_1'] = dataSub[i][4] || '';
        answers['refleksi_akhir_2'] = dataSub[i][5] || '';
      }
    }
  }

  const evaluations = {};
  if (cleanNisn) {
    const sheetEval = getOrCreateSheet('TEACHER_EVALUATIONS');
    const evalRows = sheetEval.getDataRange().getValues();
    for (let i = 1; i < evalRows.length; i++) {
      if (String(evalRows[i][2] || '').trim() !== cleanNisn) continue;
      const actId = String(evalRows[i][3] || '').trim().toLowerCase();
      if (!/^act[1-5]$/.test(actId)) continue;
      evaluations[actId] = teacherEvaluationRowToObject(evalRows[i]);
    }
  }

  return { status: 'success', progress: progress, answers: answers, evaluations: evaluations };
}



/** Mengambil profil, jawaban/progres aktivitas, dan evaluasi guru dalam satu panggilan. */
function getUserStateData(id_nisn) {
  const nisn = String(id_nisn || '').trim();
  if (!nisn) return { status: 'error', message: 'ID_NISN wajib dikirim untuk sinkronisasi.' };

  try {
    const users = getOrCreateSheet('USERS').getDataRange().getValues();
    let profile = null;
    for (let i = 1; i < users.length; i++) {
      if (String(users[i][0] || '').trim() === nisn) {
        profile = {
          id_nisn: String(users[i][0] || '').trim(),
          nama: String(users[i][1] || '').trim(),
          role: String(users[i][2] || '').trim(),
          kelas: String(users[i][3] || '').trim(),
          absen: String(users[i][4] || '').trim()
        };
        break;
      }
    }
    if (!profile) return { status: 'error', message: 'Profil pengguna tidak ditemukan pada tab USERS.' };

    const studentData = getStudentData(nisn);
    if (!studentData || studentData.status !== 'success') {
      return { status: 'error', message: 'Data aktivitas murid gagal dibaca.' };
    }
    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();
    return {
      status: 'success',
      profile: profile,
      progress: studentData.progress || {},
      answers: studentData.answers || {},
      evaluations: studentData.evaluations || {},
      synced_at: new Date().toISOString()
    };
  } catch (err) {
    return { status: 'error', message: 'Sinkronisasi data gagal: ' + err.message };
  }
}

// Unggah berkas murid ke Google Drive. Spreadsheet hanya menerima URL berkas, bukan Base64.
function uploadFileToDrive(fileObj) {
  fileObj = fileObj || {};
  const nisn = String(fileObj.nisn || '').trim();
  const actId = String(fileObj.act_id || '').trim().toLowerCase();
  const fileName = String(fileObj.fileName || fileObj.nama_file || '').trim();
  const mimeType = String(fileObj.mimeType || fileObj.mime_type || '').trim();
  let base64 = String(fileObj.base64 || '').trim();
  base64 = base64.replace(/^data:[^;]+;base64,/, '');

  if (!nisn) return { status: 'error', message: 'NISN murid tidak ditemukan.' };
  if (!/^act[1-5]$/.test(actId)) return { status: 'error', message: 'ID aktivitas tidak valid.' };
  if (!fileName || !mimeType || !base64) return { status: 'error', message: 'Data berkas tidak lengkap.' };
  if (!/^image\//i.test(mimeType) && mimeType !== 'application/pdf') {
    return { status: 'error', message: 'Format berkas harus berupa foto/gambar atau PDF.' };
  }

  try {
    const bytes = Utilities.base64Decode(base64);
    // Batas defensif 8 MB untuk menghindari payload terlalu besar.
    if (bytes.length > 8 * 1024 * 1024) {
      return { status: 'error', message: 'Ukuran berkas maksimal 8 MB. Silakan kompres foto atau PDF.' };
    }
    const safeName = fileName.replace(/[\\/:*?"<>|#%{}~&]/g, '_');
    const blob = Utilities.newBlob(bytes, mimeType, safeName);
    const folderName = 'PORTAL_SRL_UPLOADS';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const file = folder.createFile(blob).setName(`${nisn}_${actId}_${timestamp}_${safeName}`);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      status: 'success',
      message: 'Berkas berhasil diunggah ke Google Drive.',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      url: file.getUrl(),
      fileName: file.getName(),
      nisn: nisn,
      act_id: actId
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal mengunggah berkas ke Google Drive: ' + err.message };
  }
}

// Simpan atau perbarui bukti aktivitas murid pada STUDENT_SUBMISSIONS.
// Untuk act1-act5, pencocokan menggunakan kombinasi NISN + act_id.
function saveStudentSubmission(p) {
  p = p || {};
  const nisn = String(p.nisn || '').trim();
  const actId = String(p.act_id || '').trim().toLowerCase();
  const allowedActivities = ['act1', 'act2', 'act3', 'act4', 'act5'];

  if (!nisn) return { status: 'error', message: 'ID/NISN murid tidak ditemukan. Silakan login kembali.' };
  if (!actId) return { status: 'error', message: 'Kode aktivitas wajib diisi.' };
  if (!allowedActivities.includes(actId) && !['stimulasi', 'refleksi_akhir'].includes(actId)) {
    return { status: 'error', message: 'Kode aktivitas tidak valid.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOrCreateSheet('STUDENT_SUBMISSIONS');
    const lastRow = sheet.getLastRow();
    const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 11).getValues() : [];
    const timestamp = new Date();
    const subId = `SUB_${nisn}_${actId.toUpperCase()}`;
    let rowIndex = -1;

    // Prioritas utama: kombinasi nisn + act_id, bukan hanya ID baris.
    for (let i = 0; i < data.length; i++) {
      const rowNisn = String(data[i][2] || '').trim();
      const rowActId = String(data[i][3] || '').trim().toLowerCase();
      if (rowNisn === nisn && rowActId === actId) {
        rowIndex = i + 2; // Data dimulai dari baris 2.
        break;
      }
    }

    const answers = (p.answers && typeof p.answers === 'object') ? p.answers : {};
    const firstNonEmpty = (values) => {
      for (const value of values) {
        if (value !== undefined && value !== null && String(value).trim() !== '') return value;
      }
      return '';
    };

    const jawabanTeks1 = firstNonEmpty([
      p.jawaban_teks_1,
      answers[`${actId}_jawaban_1`], answers[`${actId}_identifikasi`],
      answers[`${actId}_dampak`], answers[`${actId}_organik`],
      answers[`${actId}_rancangan`], answers[`${actId}_kebiasaan`],
      answers[`${actId}_lkpd_url`]
    ]);
    const jawabanTeks2 = firstNonEmpty([
      p.jawaban_teks_2,
      answers[`${actId}_jawaban_2`], answers[`${actId}_anorganik`],
      answers[`${actId}_photo_url`]
    ]);
    const driveEvidence = {
      lkpd_url: String(p.lkpd_drive_url || answers[`${actId}_lkpd_url`] || answers[`${actId}_lkpd_drive_url`] || ''),
      photo_url: String(p.photo_drive_url || answers[`${actId}_photo_url`] || answers[`${actId}_photo_drive_url`] || '')
    };
    const driveEvidenceJson = (driveEvidence.lkpd_url || driveEvidence.photo_url) ? JSON.stringify(driveEvidence) : '';
    const linkPadlet = firstNonEmpty([
      driveEvidenceJson, p.link_karya_padlet, answers[`${actId}_padlet`],
      answers[`${actId}_video_url`], answers[`${actId}_link_video`]
    ]);

    // Simpan semua butir peer/self/refleksi aktivitas sebagai JSON string.
    const peerAssessment = p.peer_assessment !== undefined
      ? (typeof p.peer_assessment === 'string' ? p.peer_assessment : JSON.stringify(p.peer_assessment))
      : JSON.stringify(Object.fromEntries(Object.entries(answers).filter(([key]) =>
          key.startsWith(`${actId}_peer_`) || key === `${actId}_peer_target` || key === `${actId}_peer`
        )));
    const selfAssessment = p.self_assessment !== undefined
      ? (typeof p.self_assessment === 'string' ? p.self_assessment : JSON.stringify(p.self_assessment))
      : JSON.stringify(Object.fromEntries(Object.entries(answers).filter(([key]) =>
          key.startsWith(`${actId}_self_`) || key === `${actId}_self`
        )));
    const refleksiMurid = p.refleksi_murid !== undefined
      ? (typeof p.refleksi_murid === 'string' ? p.refleksi_murid : JSON.stringify(p.refleksi_murid))
      : JSON.stringify(Object.fromEntries(Object.entries(answers).filter(([key]) =>
          key.startsWith(`${actId}_refleksi`) || key.startsWith(`${actId}_kebiasaan_lanjut`) || key === `${actId}_refleksi`
        )));

    const statusMisi = String(p.status_misi || 'selesai').toLowerCase() === 'proses' ? 'proses' : 'selesai';
    const rowValues = [
      subId, timestamp, nisn, actId,
      jawabanTeks1, jawabanTeks2, linkPadlet,
      peerAssessment, selfAssessment, refleksiMurid, statusMisi
    ];

    let operation;
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      operation = 'updated';
    } else {
      sheet.appendRow(rowValues);
      operation = 'inserted';
    }

    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();

    // Buat notifikasi hanya setelah kiriman murid berhasil disimpan ke Sheets.
    try {
      createStudentSubmissionNotification_(nisn, actId, statusMisi, operation);
    } catch (notificationError) {
      // Kegagalan notifikasi tidak membatalkan kiriman tugas yang sudah tersimpan.
      console.error('Notifikasi kiriman murid gagal dibuat: ' + notificationError.message);
    }

    // Hitung progres Aktivitas 1-5 setelah penyimpanan berhasil.
    const freshLastRow = sheet.getLastRow();
    const freshData = freshLastRow >= 2 ? sheet.getRange(2, 1, freshLastRow - 1, 11).getValues() : [];
    const progress = { act1: false, act2: false, act3: false, act4: false, act5: false };
    freshData.forEach(row => {
      if (String(row[2] || '').trim() === nisn) {
        const id = String(row[3] || '').trim().toLowerCase();
        if (allowedActivities.includes(id)) progress[id] = String(row[10] || '').trim().toLowerCase() === 'selesai';
      }
    });
    const completedCount = allowedActivities.filter(id => progress[id]).length;
    const progressPercent = Math.round((completedCount / allowedActivities.length) * 100);

    return {
      status: 'success',
      message: statusMisi === 'selesai' ? 'Aktivitas berhasil disimpan.' : 'Progres aktivitas berhasil diperbarui.',
      operation: operation,
      submission_id: subId,
      nisn: nisn,
      act_id: actId,
      status_misi: statusMisi,
      progress: progress,
      completed_activities: completedCount,
      total_activities: allowedActivities.length,
      progress_percent: progressPercent
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal menyimpan aktivitas: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// Tambah data murid ke USERS dengan validasi NISN unik.
// Kolom USERS: ID_NISN, NAMA, ROLE, KELAS, ABSEN, PASSWORD.
function addStudent(studentData) {
  const data = studentData || {};
  const nisn = String(data.nisn || data.ID_NISN || '').trim();
  const nama = String(data.nama || data.nama_lengkap || data.NAMA || '').trim();
  const kelas = String(data.kelas || data.KELAS || '').trim();
  const absen = String(data.absen || data.nomor_absen || data.ABSEN || '').trim();

  if (!nisn || !nama || !kelas || !absen) {
    return { status: 'error', message: 'NISN, nama lengkap, kelas, dan nomor absen wajib diisi.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getOrCreateSheet('USERS');
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      const exists = existingIds.some(row => String(row[0] || '').trim() === nisn);
      if (exists) return { status: 'error', message: 'NISN sudah terdaftar.' };
    }

    // Password awal disamakan dengan NISN sesuai ketentuan.
    sheet.appendRow([nisn, nama, 'Murid', kelas, absen, nisn]);
    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();
    return {
      status: 'success',
      message: 'Data murid berhasil ditambahkan.',
      student: { nisn: nisn, nama: nama, role: 'Murid', kelas: kelas, absen: absen }
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal menambahkan murid: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Ambil Seluruh Data Rekap untuk Dashboard Guru
function getTeacherData() {
  const sheetUsers = getOrCreateSheet('USERS');
  const sheetSub = getOrCreateSheet('STUDENT_SUBMISSIONS');
  const sheetEval = getOrCreateSheet('TEACHER_EVALUATIONS');

  const users = sheetUsers.getDataRange().getValues();
  const subs = sheetSub.getDataRange().getValues();
  const evals = sheetEval.getDataRange().getValues();

  const students = [];
  for (let i = 1; i < users.length; i++) {
    const role = String(users[i][2] || '').trim().toLowerCase();
    if (role === 'murid' || role === 'siswa') {
      students.push({
        nisn: String(users[i][0]).trim(),
        nama: String(users[i][1]).trim(),
        kelas: String(users[i][3]).trim(),
        absen: String(users[i][4] || '').trim()
      });
    }
  }

  const submissions = {};
  for (let i = 1; i < subs.length; i++) {
    const nisn = String(subs[i][2]).trim();
    const actId = String(subs[i][3]).toLowerCase();
    if (!submissions[nisn]) submissions[nisn] = {};
    submissions[nisn][actId] = {
      jawaban_1: subs[i][4],
      jawaban_2: subs[i][5],
      padlet: subs[i][6],
      peer: subs[i][7],
      self: subs[i][8],
      refleksi: subs[i][9],
      status: subs[i][10]
    };
  }

  const evaluations = {};
  for (let i = 1; i < evals.length; i++) {
    const nisn = String(evals[i][2]).trim();
    const actId = String(evals[i][3]).toLowerCase();
    if (!evaluations[nisn]) evaluations[nisn] = {};
    evaluations[nisn][actId] = teacherEvaluationRowToObject(evals[i]);
  }

  return {
    status: 'success',
    students: students,
    submissions: submissions,
    evaluations: evaluations
  };
}

// Parse nilai yang tersimpan sebagai JSON, tetapi tetap mendukung nilai teks biasa.
function parseEvaluationValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value !== 'string') return value;
  try { return JSON.parse(value); } catch (e) { return value; }
}

function teacherEvaluationRowToObject(row) {
  return {
    eval_id: row[0] || '',
    timestamp: row[1] instanceof Date ? row[1].toISOString() : (row[1] || ''),
    nisn: String(row[2] || '').trim(),
    act_id: String(row[3] || '').trim().toLowerCase(),
    rubrik_penilaian: parseEvaluationValue(row[4]),
    observasi_guru: row[5] || '',
    apk_penalaran: parseEvaluationValue(row[6]),
    apk_kreativitas: parseEvaluationValue(row[7]),
    apk_kolaborasi: parseEvaluationValue(row[8]),
    feedback_guru: row[9] || '',
    tindak_lanjut: row[10] || ''
  };
}

// Simpan evaluasi guru. Kunci uniknya adalah kombinasi NISN + act_id.
function saveTeacherEvaluation(p) {
  p = p || {};
  const nisn = String(p.nisn || '').trim();
  const actId = String(p.act_id || '').trim().toLowerCase();
  if (!nisn) return { status: 'error', message: 'NISN murid wajib diisi.' };
  if (!/^act[1-5]$/.test(actId)) return { status: 'error', message: 'Kode aktivitas harus act1 sampai act5.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getOrCreateSheet('TEACHER_EVALUATIONS');
    const lastRow = sheet.getLastRow();
    const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 11).getValues() : [];
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowNisn = String(rows[i][2] || '').trim();
      const rowAct = String(rows[i][3] || '').trim().toLowerCase();
      if (rowNisn === nisn && rowAct === actId) { rowIndex = i + 2; break; }
    }

    const evalId = `EVAL_${nisn}_${actId.toUpperCase()}`;
    const jsonOrText = value => {
      if (value === null || value === undefined || value === '') return '';
      if (typeof value === 'string') return value;
      try { return JSON.stringify(value); } catch (e) { return String(value); }
    };
    const rubrik = p.rubrik_penilaian !== undefined ? p.rubrik_penilaian : (p.ratings || {});
    const apkPenalaran = p.apk_penalaran !== undefined ? p.apk_penalaran : (p.apk_penalaran_kritis || '');
    const apkKreativitas = p.apk_kreativitas !== undefined ? p.apk_kreativitas : '';
    const apkKolaborasi = p.apk_kolaborasi !== undefined ? p.apk_kolaborasi : '';
    const observasi = p.observasi_guru !== undefined ? p.observasi_guru : (p.catatanGuru || '');
    const feedback = p.feedback_guru !== undefined ? p.feedback_guru : (p.catatanGuru || '');
    const tindakLanjut = p.tindak_lanjut !== undefined ? p.tindak_lanjut : (p.tindakLanjut || '');
    const rowValues = [
      evalId, new Date(), nisn, actId,
      jsonOrText(rubrik), String(observasi || ''),
      jsonOrText(apkPenalaran), jsonOrText(apkKreativitas), jsonOrText(apkKolaborasi),
      String(feedback || ''), String(tindakLanjut || '')
    ];

    let operation;
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, 11).setValues([rowValues]);
      operation = 'updated';
    } else {
      sheet.appendRow(rowValues);
      rowIndex = sheet.getLastRow();
      operation = 'inserted';
    }
    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();
    return {
      status: 'success', message: operation === 'updated' ? 'Evaluasi guru berhasil diperbarui.' : 'Evaluasi guru berhasil disimpan.',
      operation: operation, eval_id: evalId, nisn: nisn, act_id: actId,
      evaluation: teacherEvaluationRowToObject(sheet.getRange(rowIndex, 1, 1, 11).getValues()[0])
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal menyimpan evaluasi guru: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Menjembatani form Rekap/Penilaian APK pada Portal Guru ke tab TEACHER_EVALUATIONS.
// Rubrik disimpan per aktivitas; nilai subdimensi ditampilkan di Portal Murid berdasarkan data ini.
function saveTeacherActivityAssessment(p) {
  p = p || {};
  const nisn = String(p.nisn || '').trim();
  const actId = String(p.act_id || '').trim().toLowerCase();
  const ratings = p.ratings && typeof p.ratings === 'object' ? p.ratings : {};
  if (!nisn) return { status: 'error', message: 'NISN murid wajib diisi.' };
  if (!/^act[1-5]$/.test(actId)) return { status: 'error', message: 'Aktivitas harus act1 sampai act5.' };
  const allowed = ['Berkembang', 'Cakap', 'Mahir'];
  const cleanRatings = {};
  Object.keys(ratings).forEach(key => {
    const level = String(ratings[key] || '').trim();
    if (key && allowed.includes(level)) cleanRatings[String(key).trim()] = level;
  });
  if (!Object.keys(cleanRatings).length) return { status: 'error', message: 'Pilih capaian APK yang valid sebelum menyimpan.' };

  const groups = { penalaran: {}, kreativitas: {}, kolaborasi: {} };
  const dimensionMap = {
    'Penyampaian Argumentasi': 'penalaran', 'Pengambilan Keputusan': 'penalaran', 'Penyelesaian Masalah': 'penalaran',
    'Gagasan Baru': 'kreativitas', 'Fleksibilitas Berpikir': 'kreativitas', 'Karya': 'kreativitas',
    'Peduli': 'kolaborasi', 'Berbagi': 'kolaborasi', 'Kerja Sama': 'kolaborasi'
  };
  Object.keys(cleanRatings).forEach(key => {
    const dimension = dimensionMap[key];
    if (dimension) groups[dimension][key] = cleanRatings[key];
  });
  return saveTeacherEvaluation({
    nisn: nisn,
    act_id: actId,
    rubrik_penilaian: cleanRatings,
    apk_penalaran: groups.penalaran,
    apk_kreativitas: groups.kreativitas,
    apk_kolaborasi: groups.kolaborasi,
    observasi_guru: String(p.catatanGuru || p.observasi_guru || ''),
    feedback_guru: String(p.catatanGuru || p.feedback_guru || ''),
    tindak_lanjut: String(p.tindakLanjut || p.tindak_lanjut || '')
  });
}

// Mengambil capaian sembilan subdimensi untuk satu murid dari evaluasi guru yang tersimpan.
// Jika subdimensi dinilai pada beberapa aktivitas, data dari evaluasi terbaru yang dipakai.
function getStudentAPKProfile(nisn) {
  const cleanNisn = String(nisn || '').trim();
  if (!cleanNisn) return { status: 'error', message: 'NISN murid wajib diisi.' };
  const validLevels = ['Berkembang', 'Cakap', 'Mahir'];
  const labels = [
    'Penyampaian Argumentasi', 'Pengambilan Keputusan', 'Penyelesaian Masalah',
    'Gagasan Baru', 'Fleksibilitas Berpikir', 'Karya', 'Peduli', 'Berbagi', 'Kerja Sama'
  ];
  const sheet = getOrCreateSheet('TEACHER_EVALUATIONS');
  const rows = sheet.getDataRange().getValues();
  const selected = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2] || '').trim() !== cleanNisn) continue;
    const actId = String(rows[i][3] || '').trim().toLowerCase();
    if (!/^act[1-5]$/.test(actId)) continue;
    const timestampRaw = rows[i][1];
    const timestamp = timestampRaw instanceof Date ? timestampRaw.getTime() : (Date.parse(timestampRaw) || 0);
    const rubric = parseEvaluationValue(rows[i][4]);
    const candidates = [];
    if (rubric && typeof rubric === 'object' && !Array.isArray(rubric)) {
      Object.keys(rubric).forEach(key => candidates.push([String(key).trim(), rubric[key]]));
    }
    // Mendukung data lama yang menyimpan kategori pada kolom APK per dimensi.
    [[rows[i][6]], [rows[i][7]], [rows[i][8]]].forEach(col => {
      const parsed = parseEvaluationValue(col[0]);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) Object.keys(parsed).forEach(key => candidates.push([String(key).trim(), parsed[key]]));
    });
    candidates.forEach(([label, level]) => {
      const cleanLevel = String(level || '').trim();
      if (!labels.includes(label) || !validLevels.includes(cleanLevel)) return;
      if (!selected[label] || timestamp >= selected[label]._timestamp) {
        selected[label] = { level: cleanLevel, act_id: actId, updatedAt: timestampRaw instanceof Date ? timestampRaw.toISOString() : String(timestampRaw || ''), _timestamp: timestamp };
      }
    });
  }
  const subdimensions = {};
  labels.forEach(label => {
    if (selected[label]) {
      const value = selected[label];
      subdimensions[label] = { level: value.level, act_id: value.act_id, updatedAt: value.updatedAt };
    }
  });
  const timestamps = Object.keys(subdimensions).map(key => Date.parse(subdimensions[key].updatedAt) || 0).filter(Boolean);
  return { status: 'success', nisn: cleanNisn, subdimensions: subdimensions, assessedCount: Object.keys(subdimensions).length, updatedAt: timestamps.length ? new Date(Math.max.apply(null, timestamps)).toISOString() : '' };
}

// Pembaca evaluasi untuk satu murid dan satu aktivitas. Tanpa act_id, mengembalikan semua act1-act5.
function getTeacherEvaluation(nisn, actId) {
  const cleanNisn = String(nisn || '').trim();
  const cleanActId = String(actId || '').trim().toLowerCase();
  if (!cleanNisn) return { status: 'error', message: 'NISN murid wajib diisi.' };
  if (cleanActId && !/^act[1-5]$/.test(cleanActId)) return { status: 'error', message: 'Kode aktivitas harus act1 sampai act5.' };

  const sheet = getOrCreateSheet('TEACHER_EVALUATIONS');
  const rows = sheet.getDataRange().getValues();
  const evaluations = {};
  for (let i = 1; i < rows.length; i++) {
    const rowNisn = String(rows[i][2] || '').trim();
    const rowActId = String(rows[i][3] || '').trim().toLowerCase();
    if (rowNisn !== cleanNisn || !/^act[1-5]$/.test(rowActId)) continue;
    if (cleanActId && rowActId !== cleanActId) continue;
    evaluations[rowActId] = teacherEvaluationRowToObject(rows[i]);
  }
  if (cleanActId) {
    return { status: 'success', nisn: cleanNisn, act_id: cleanActId, evaluation: evaluations[cleanActId] || null, found: !!evaluations[cleanActId] };
  }
  return { status: 'success', nisn: cleanNisn, evaluations: evaluations, found: Object.keys(evaluations).length > 0 };
}


/** Menyimpan status akses aktivitas. status=true berarti aktif/terbuka. */
function saveRestrictionStatus(aktivitasId, status) {
  const id = String(aktivitasId || '').trim().toLowerCase();
  if (!/^act[1-5]$/.test(id)) return {status:'error', message:'ID aktivitas tidak valid.'};
  const active = status === true || String(status).toLowerCase() === 'true' || String(status).toLowerCase() === 'on' || String(status).toLowerCase() === 'aktif';
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getOrCreateSheet('Pengaturan_Aktivitas');
    const rows = sh.getDataRange().getValues();
    let row = -1;
    for (let i=1; i<rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === id) { row=i+1; break; }
    const record = [id, active ? 'ON' : 'OFF', new Date()];
    if (row > 0) sh.getRange(row,1,1,3).setValues([record]); else sh.appendRow(record);
    // Sinkronkan juga objek teacherSettings yang sudah digunakan pengecekan akses murid.
    const settingsSheet = getOrCreateSheet('Pengaturan');
    const settingRows = settingsSheet.getDataRange().getValues();
    let settingsRow = -1, settingsObject = {};
    for (let i=1; i<settingRows.length; i++) {
      if (String(settingRows[i][0] || '').trim() === 'teacherSettings') {
        settingsRow = i + 1;
        try { settingsObject = JSON.parse(String(settingRows[i][1] || '{}')); } catch(e) { settingsObject = {}; }
        break;
      }
    }
    settingsObject = settingsObject && typeof settingsObject === 'object' ? settingsObject : {};
    settingsObject.unlockedActivities = settingsObject.unlockedActivities || {};
    settingsObject.unlockedActivities[id] = active;
    const settingRecord = ['teacherSettings', JSON.stringify(settingsObject), new Date()];
    if (settingsRow > 0) settingsSheet.getRange(settingsRow,1,1,3).setValues([settingRecord]);
    else settingsSheet.appendRow(settingRecord);
    SpreadsheetApp.flush();
    invalidateInitialAppDataCache_();
    return {status:'success', aktivitasId:id, active:active, statusText:active?'ON':'OFF'};
  } catch (err) {
    return {status:'error', message:'Gagal menyimpan status aktivitas: '+err.message};
  } finally { try { lock.releaseLock(); } catch(e) {} }
}

/** Membaca status akses aktivitas; default ON untuk aktivitas yang belum diatur. */
function getRestrictionStatus() {
  try {
    const sh = getOrCreateSheet('Pengaturan_Aktivitas');
    const rows = sh.getDataRange().getValues();
    const statuses = {act1:true,act2:true,act3:true,act4:true,act5:true};
    for (let i=1; i<rows.length; i++) {
      const id = String(rows[i][0] || '').trim().toLowerCase();
      if (/^act[1-5]$/.test(id)) statuses[id] = String(rows[i][1] || '').trim().toUpperCase() !== 'OFF';
    }
    return {status:'success', restrictions:statuses};
  } catch(err) { return {status:'error', message:'Gagal membaca status akses: '+err.message}; }
}

/** Ringkasan beranda dihitung dari baris data aktual, bukan angka contoh frontend. */
function getDashboardStats() {
  try {
    const ss = getSS();
    const readRows = (name) => {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return {headers:[], rows:[]};
      const v = sh.getDataRange().getValues();
      return {headers:v[0].map(x=>String(x||'').trim().toUpperCase()), rows:v.slice(1).filter(r=>r.some(c=>c!=='' && c!==null))};
    };
    const users = readRows('Data_Murid').rows.length ? readRows('Data_Murid') : readRows('USERS');
    const userHeaders = users.headers;
    let studentCount = 0;
    if (userHeaders.length) {
      const roleIndex = userHeaders.findIndex(h=>['ROLE','PERAN','JENIS_AKUN'].includes(h));
      studentCount = users.rows.filter(r => roleIndex < 0 || /murid|siswa|student/i.test(String(r[roleIndex]||''))).length;
    }
    const submissions = ['Bukti_Karya','Aktivitas','STUDENT_SUBMISSIONS'].map(readRows);
    const evidenceSource = submissions.find(s=>s.rows.length) || submissions[0];
    const evidenceCount = evidenceSource.rows.length;
    const evalData = ['Penilaian_APK','PENILAIAN_APK','TEACHER_EVALUATIONS','Penilaian'].map(readRows).find(s=>s.rows.length) || {headers:[],rows:[]};
    const scoreIndexes = evalData.headers.map((h,i)=>({h,i})).filter(x=>/APK|NILAI|SKOR|CAPAIAN/.test(x.h));
    let scoreSum=0, scoreCount=0;
    evalData.rows.forEach(r=>scoreIndexes.forEach(x=>{
      const n = Number(String(r[x.i] == null ? '' : r[x.i]).replace('%','').replace(',','.'));
      if (String(r[x.i] == null ? '' : r[x.i]).trim() !== '' && Number.isFinite(n) && n >= 0 && n <= 100) { scoreSum += n; scoreCount++; }
    }));
    const settingsRes = getPengaturan();
    const settings = settingsRes && settingsRes.settings ? settingsRes.settings : {};
    const teacherSettingsForTopic = settings.teacherSettings && typeof settings.teacherSettings === 'object' ? settings.teacherSettings : {};
    let selectedTopic = settings.selectedTopik || settings.selectedTopic || teacherSettingsForTopic.selectedTopik || 'Topik 2: Aku Bisa Menjaga Lingkungan';
    const restrictionRes = getRestrictionStatus();
    const restrictions = restrictionRes && restrictionRes.restrictions ? restrictionRes.restrictions : {act1:true,act2:true,act3:true,act4:true,act5:true};
    const subData = readRows('STUDENT_SUBMISSIONS');
    const nisnIdx = subData.headers.findIndex(h=>['NISN','ID_NISN'].includes(h));
    const actIdx = subData.headers.findIndex(h=>['ACT_ID','AKTIVITAS_ID'].includes(h));
    const progressByActivity = {};
    ['act1','act2','act3','act4','act5'].forEach(id=>{
      const matching = subData.rows.filter(r=>String(r[actIdx]||'').toLowerCase()===id);
      const uniqueStudents = new Set(matching.map(r=>String(r[nisnIdx]||'').trim()).filter(Boolean)).size;
      progressByActivity[id] = {submitted:matching.length, students:uniqueStudents, percent:studentCount ? Math.round(uniqueStudents/studentCount*100) : 0};
    });
    return {status:'success', stats:{studentCount:studentCount, activeTopic:selectedTopic, evidenceCount:evidenceCount, averageApk:scoreCount?Math.round(scoreSum/scoreCount):0, progressByActivity:progressByActivity, restrictions:restrictions}};
  } catch(err) { return {status:'error', message:'Gagal menghitung ringkasan beranda: '+err.message}; }
}


/* ============================================================================
 * NOTIFIKASI TERPADU PORTAL MURID -> PORTAL GURU
 * Tab NOTIFICATIONS dibuat otomatis tanpa mengubah tab/data yang sudah ada.
 * ========================================================================== */
const SRL_NOTIFICATION_HEADERS_ = [
  'NOTIFICATION_ID', 'TIMESTAMP', 'NISN', 'NAMA_MURID', 'KELAS',
  'ACT_ID', 'JUDUL', 'PESAN', 'STATUS_BACA', 'JENIS', 'OPERATION'
];

function getNotificationsSheet_() {
  const ss = (typeof getSS === 'function') ? getSS() : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet PORTAL-SRL tidak ditemukan.');
  let sh = ss.getSheetByName('NOTIFICATIONS');
  if (!sh) sh = ss.insertSheet('NOTIFICATIONS');
  if (sh.getLastRow() === 0) {
    sh.appendRow(SRL_NOTIFICATION_HEADERS_);
  } else {
    // Tambahkan header baru di ujung bila tab lama sudah ada, tanpa menghapus data.
    const existing = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .getDisplayValues()[0].map(v => String(v || '').trim());
    SRL_NOTIFICATION_HEADERS_.forEach(header => {
      if (!existing.includes(header)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
        existing.push(header);
      }
    });
  }
  return sh;
}

function createStudentSubmissionNotification_(nisn, actId, statusMisi, operation) {
  const users = getOrCreateSheet('USERS').getDataRange().getValues();
  let studentName = 'Murid';
  let studentClass = '';
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0] || '').trim() === String(nisn).trim()) {
      studentName = String(users[i][1] || 'Murid').trim();
      studentClass = String(users[i][3] || '').trim();
      break;
    }
  }

  const activityNames = {
    act1: 'Aktivitas 1',
    act2: 'Aktivitas 2',
    act3: 'Aktivitas 3',
    act4: 'Aktivitas 4',
    act5: 'Aktivitas 5',
    stimulasi: 'Stimulasi',
    refleksi_akhir: 'Refleksi Akhir'
  };
  const activityName = activityNames[actId] || actId;
  const sh = getNotificationsSheet_();
  const id = 'NTF_' + Utilities.getUuid();
  const actionText = operation === 'updated' ? 'mengirim ulang' : 'mengirim';
  const statusText = statusMisi === 'selesai' ? 'Tugas' : 'Progres';
  sh.appendRow([
    id, new Date(), String(nisn), studentName, studentClass,
    String(actId), `${statusText} ${activityName} baru`,
    `${studentName} (${studentClass || 'kelas belum tercatat'}) ${actionText} ${activityName}. Silakan periksa bukti dan jawaban murid.`,
    'BELUM_DIBACA', 'KIRIMAN_MURID', operation
  ]);
  SpreadsheetApp.flush();
  return id;
}

function getTeacherNotifications(payload) {
  try {
    payload = payload || {};
    const sh = getNotificationsSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { status: 'success', notifications: [], unreadCount: 0 };

    const headers = values[0].map(v => String(v || '').trim());
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);
    const notifications = [];
    const requestedClass = String(payload.teacherClass || '').trim();
    const normalizeClass = value => String(value || '').toLowerCase().replace(/kelas/g, '').replace(/sd/g, '').replace(/iv/g, '4').replace(/v/g, '5').replace(/vi/g, '6').replace(/[^a-z0-9]/g, '');
    for (let i = values.length - 1; i >= 1 && notifications.length < 60; i--) {
      const row = values[i];
      const rowClass = String(row[idx.KELAS] || '').trim();
      if (requestedClass && normalizeClass(rowClass) !== normalizeClass(requestedClass)) continue;
      const ts = row[idx.TIMESTAMP];
      notifications.push({
        id: String(row[idx.NOTIFICATION_ID] || ''),
        timestamp: ts instanceof Date ? ts.toISOString() : String(ts || ''),
        nisn: String(row[idx.NISN] || ''),
        studentName: String(row[idx.NAMA_MURID] || 'Murid'),
        kelas: String(row[idx.KELAS] || ''),
        actId: String(row[idx.ACT_ID] || ''),
        title: String(row[idx.JUDUL] || 'Notifikasi PORTAL-SRL'),
        message: String(row[idx.PESAN] || ''),
        read: String(row[idx.STATUS_BACA] || '').toUpperCase() === 'DIBACA',
        type: String(row[idx.JENIS] || 'KIRIMAN_MURID')
      });
    }
    const unreadCount = notifications.filter(n => !n.read).length;
    return { status: 'success', notifications, unreadCount };
  } catch (err) {
    return { status: 'error', message: 'Gagal mengambil notifikasi: ' + err.message, notifications: [], unreadCount: 0 };
  }
}

function markTeacherNotificationRead(payload) {
  try {
    payload = payload || {};
    const id = String(payload.notificationId || payload.id || '').trim();
    if (!id) return { status: 'error', message: 'ID notifikasi tidak ditemukan.' };
    const sh = getNotificationsSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(v => String(v || '').trim());
    const idCol = headers.indexOf('NOTIFICATION_ID');
    const readCol = headers.indexOf('STATUS_BACA');
    if (idCol < 0 || readCol < 0) return { status: 'error', message: 'Kolom notifikasi tidak tersedia.' };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol] || '') === id) {
        sh.getRange(i + 1, readCol + 1).setValue('DIBACA');
        return { status: 'success', message: 'Notifikasi ditandai sudah dibaca.' };
      }
    }
    return { status: 'error', message: 'Notifikasi tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: 'Gagal memperbarui notifikasi: ' + err.message };
  }
}


/** Sesi guru singkat untuk melindungi penyimpanan umpan balik. */
function createTeacherSession(nip) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('PORTAL_SRL_TEACHER_' + token, String(nip), 21600);
  return token;
}

function requireTeacherSession(payload) {
  const token = String((payload && payload.teacherToken) || '').trim();
  if (!token) throw new Error('Sesi guru tidak ditemukan. Silakan login kembali.');
  const nip = CacheService.getScriptCache().get('PORTAL_SRL_TEACHER_' + token);
  if (!nip) throw new Error('Sesi guru telah berakhir. Silakan login kembali.');
  const sheet = getOrCreateSheet('USERS');
  const values = sheet.getDataRange().getDisplayValues();
  const headers = (values[0] || []).map(v => String(v || '').trim().toUpperCase());
  const idIdx = headers.indexOf('ID_NISN');
  const nameIdx = headers.indexOf('NAMA');
  const roleIdx = headers.indexOf('ROLE');
  const classIdx = headers.indexOf('KELAS');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx] || '').trim() !== String(nip).trim()) continue;
    const role = String(values[i][roleIdx] || '').trim().toLowerCase();
    if (role !== 'guru' && role !== 'teacher') continue;
    return {
      nip: String(values[i][idIdx] || '').trim(),
      nama: String(values[i][nameIdx] || '').trim(),
      kelas: String(values[i][classIdx] || '').trim()
    };
  }
  throw new Error('Akun guru tidak ditemukan atau tidak aktif. Silakan login kembali.');
}

/** Menyimpan umpan balik guru secara permanen di Google Sheets. */
function saveTeacherFeedback(payload) {
  const teacher = requireTeacherSession(payload || {});
  const studentNisn = String(payload && payload.studentNisn || '').trim();
  const studentName = String(payload && payload.studentName || '').trim().slice(0, 160);
  const studentClass = String(payload && payload.studentClass || '').trim().slice(0, 80);
  const actId = String(payload && payload.actId || '').trim();
  const type = String(payload && payload.type || 'Apresiasi').trim().slice(0, 80);
  const message = String(payload && payload.message || '').trim().slice(0, 500);
  if (!studentNisn || !message) return {status:'error', message:'Pilih murid dan isi umpan balik wajib diisi.'};
  if (!/^act[1-5]$/.test(actId)) return {status:'error', message:'Aktivitas umpan balik tidak valid.'};

  const users = getOrCreateSheet('USERS');
  const rows = users.getDataRange().getDisplayValues();
  if (!rows.length) return {status:'error', message:'Data pengguna belum tersedia.'};
  const headers = rows[0].map(v => String(v || '').trim().toUpperCase());
  const idIdx = headers.indexOf('ID_NISN');
  const nameIdx = headers.indexOf('NAMA');
  const roleIdx = headers.indexOf('ROLE');
  const classIdx = headers.indexOf('KELAS');
  if (idIdx < 0 || roleIdx < 0) return {status:'error', message:'Struktur tab USERS belum sesuai.'};

  let student = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx] || '').trim() !== studentNisn) continue;
    const role = String(rows[i][roleIdx] || '').trim().toLowerCase();
    if (!['murid','siswa','student'].includes(role)) continue;
    student = {nisn:studentNisn,nama:String(rows[i][nameIdx] || studentName).trim(),kelas:String(rows[i][classIdx] || studentClass).trim()};
    break;
  }
  if (!student) return {status:'error', message:'Akun murid tidak ditemukan pada database.'};
  if (normalizeKelas(student.kelas) !== normalizeKelas(teacher.kelas)) {
    return {status:'error', message:'Guru hanya dapat mengirim umpan balik kepada murid dari kelas yang dikelolanya.'};
  }

  const sh = getOrCreateSheet('TEACHER_FEEDBACK');
  const header = ['FEEDBACK_ID','TIMESTAMP','STUDENT_NISN','STUDENT_NAME','STUDENT_CLASS','ACT_ID','TYPE','MESSAGE','TEACHER_NIP','TEACHER_NAME'];
  if (sh.getLastRow() === 0) sh.appendRow(header);
  else {
    const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),header.length)).getDisplayValues()[0];
    if (String(existing[0] || '').trim().toUpperCase() !== 'FEEDBACK_ID') return {status:'error',message:'Header tab TEACHER_FEEDBACK tidak sesuai.'};
  }
  const id = Utilities.getUuid();
  const now = new Date();
  sh.appendRow([id,now,student.nisn,student.nama,student.kelas,actId,type,message,teacher.nip,teacher.nama]);
  SpreadsheetApp.flush();
  return {status:'success',message:'Umpan balik berhasil disimpan.',feedback:{id,timestamp:now.toISOString(),nisn:student.nisn,actId,type,message,teacherName:teacher.nama}};
}

/** Mengambil hanya umpan balik yang cocok dengan ID murid yang sedang masuk. */
function getStudentFeedback(nisn) {
  const studentNisn = String(nisn || '').trim();
  if (!studentNisn) return {status:'error',message:'Akun murid belum teridentifikasi.'};
  const users = getOrCreateSheet('USERS');
  const userRows = users.getDataRange().getDisplayValues();
  const userHeaders = (userRows[0] || []).map(v=>String(v||'').trim().toUpperCase());
  const idIdx = userHeaders.indexOf('ID_NISN');
  const roleIdx = userHeaders.indexOf('ROLE');
  let isStudent = false;
  for (let i=1;i<userRows.length;i++) {
    if (String(userRows[i][idIdx]||'').trim() !== studentNisn) continue;
    isStudent = ['murid','siswa','student'].includes(String(userRows[i][roleIdx]||'').trim().toLowerCase());
    break;
  }
  if (!isStudent) return {status:'error',message:'Akun murid tidak ditemukan.'};
  const sh = getOrCreateSheet('TEACHER_FEEDBACK');
  if (sh.getLastRow() < 2) return {status:'success',feedback:[]};
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(v=>String(v||'').trim().toUpperCase());
  const idx = {}; headers.forEach((h,i)=>idx[h]=i);
  if (['STUDENT_NISN','ACT_ID','MESSAGE'].some(h=>idx[h]===undefined)) return {status:'error',message:'Struktur data umpan balik belum sesuai.'};
  const feedback=[];
  for (let i=1;i<values.length;i++) {
    const row=values[i];
    if (String(row[idx.STUDENT_NISN]||'').trim() !== studentNisn) continue;
    const dateValue=row[idx.TIMESTAMP];
    feedback.push({
      id:String(row[idx.FEEDBACK_ID]||''),
      timestamp:dateValue instanceof Date?dateValue.toISOString():String(dateValue||''),
      actId:String(row[idx.ACT_ID]||''),
      type:String(row[idx.TYPE]||'Umpan Balik'),
      message:String(row[idx.MESSAGE]||''),
      teacherName:String(row[idx.TEACHER_NAME]||'Guru')
    });
  }
  feedback.sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));
  return {status:'success',feedback};
}
