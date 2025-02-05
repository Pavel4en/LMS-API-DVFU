import axios from 'axios';
import * as XLSX from 'xlsx';

/* ======================= ОБЩИЕ НАСТРОЙКИ И ФУНКЦИИ ======================= */
const API_BASE = 'https://lms.dvfu.ru/endpoint/v1';
const TOKEN_URL = 'https://lms.dvfu.ru/oauth/token';
const COURSES_URL = `${API_BASE}/courses`;
const COURSE_TYPES_URL = `${API_BASE}/course_types`;
const FORM_URL = 'https://forms.yandex.ru/cloud/6743cc36c417f388901ebaf6/';

const CLIENT_ID = 'kPQ3YqMfpYlDr200G1gE6ItXSSRGiuK6Am8lFUR_ei4';
const CLIENT_SECRET = 'jIyrta8XZJQZjK_jv1ld-43AXZLcMn9oi98qQRQVW0M';

let accessToken = null;
let tokenExpiry = null; // в мс

// Элементы для "Курсы и потоки"
const logContainer = document.getElementById('log');
const progressDiv = document.getElementById('progress');
const fullExportTableBody = document.querySelector('#fullExportTable tbody');
const courseTypesContainer = document.getElementById('courseTypesContainer');

// Элементы для "Разделы и материалы"
const logMaterials = document.getElementById('logMaterials');
const progressMaterials = document.getElementById('progressMaterials');
const materialsDetailsTableBody = document.querySelector('#materialsDetailsTable tbody');

// Элементы для "Обратная связь"
const feedbackFileInput = document.getElementById('feedbackFileInput');
const processFeedbackBtn = document.getElementById('processFeedbackBtn'); // кнопка "Добавить ОС"
const exportFeedbackBtn = document.getElementById('exportFeedbackBtn');   // кнопка "Выгрузить результат"
const progressFeedback = document.getElementById('progressFeedback');
const logFeedback = document.getElementById('logFeedback');

// NEW: Элемент для загрузки файла с course_id из фильтра
const courseIdFileInput = document.getElementById('courseIdFileInput');

// При выборе файла с course_id сразу обрабатываем его
courseIdFileInput.addEventListener('change', () => {
  if (courseIdFileInput.files && courseIdFileInput.files.length > 0) {
    processCourseIdFilterFile(courseIdFileInput.files[0]);
  }
});

const { ipcRenderer } = require('electron');

ipcRenderer.on('update-message', (event, message) => {
  console.log('Update message:', message);
  const updateDiv = document.getElementById('updateMessage');
  if (updateDiv) {
    updateDiv.textContent = message;
  }
});

ipcRenderer.on('update-progress', (event, percent) => {
  console.log('Update progress:', percent, '%');
  const updateDiv = document.getElementById('updateMessage');
  if (updateDiv) {
    updateDiv.textContent = `Скачивание обновления: ${percent}%`;
  }
});



// Глобальные данные
window.fullExportData = []; // для "Курсы и потоки"
window.filterOptions = { startDate: null, endDate: null, courseTypes: [], courseIds: [] };

window.materialsDataDetails = []; // для "Разделы и материалы"
window.filterOptionsMaterials = { startDate: null, endDate: null, courseTypes: [], courseIds: [] };

window.feedbackData = [];    // исходные данные из файла (сгенерированные ссылки)
window.feedbackResults = []; // результаты создания разделов и материалов

let currentFilterTarget = 'courses'; // для фильтров

/* ======================= ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ФОРМАТИРОВАНИЯ ВРЕМЕНИ ======================= */
function formatTime(seconds) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return `${h} ч ${m} мин ${s} сек`;
}

/* ======================= ФУНКЦИИ ЛОГИРОВАНИЯ И ПРОГРЕССА ======================= */
function addLog(message) {
  const time = new Date().toLocaleTimeString();
  const p = document.createElement('p');
  p.textContent = `[${time}] ${message}`;
  logContainer.appendChild(p);
  logContainer.scrollTop = logContainer.scrollHeight;
  if (window.electronAPI?.logMessage) {
    window.electronAPI.logMessage(message);
  }
}

function addLogMaterials(message) {
  const time = new Date().toLocaleTimeString();
  const p = document.createElement('p');
  p.textContent = `[${time}] ${message}`;
  logMaterials.appendChild(p);
  logMaterials.scrollTop = logMaterials.scrollHeight;
  if (window.electronAPI?.logMessage) {
    window.electronAPI.logMessage(message);
  }
}

function addLogFeedback(message) {
  const time = new Date().toLocaleTimeString();
  const p = document.createElement('p');
  p.textContent = `[${time}] ${message}`;
  logFeedback.appendChild(p);
  logFeedback.scrollTop = logFeedback.scrollHeight;
  if (window.electronAPI?.logMessage) {
    window.electronAPI.logMessage(message);
  }
}

function setProgress(message) {
  progressDiv.textContent = message;
}

function setProgressMaterials(message) {
  progressMaterials.textContent = message;
}

function setProgressFeedback(message) {
  progressFeedback.textContent = message;
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/* ======================= ФУНКЦИИ ДЛЯ ТОКЕНА ======================= */
async function getAccessToken() {
  setProgress('Запрос нового токена...');
  addLog('Запрос нового токена...');
  try {
    const response = await axios.post(TOKEN_URL, {
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: 'client_credentials'
    });
    accessToken = response.data.access_token;
    tokenExpiry = Date.now() + response.data.expires_in * 1000;
    addLog('Токен успешно получен.');
  } catch (error) {
    addLog(`Ошибка получения токена: ${error}`);
    throw error;
  }
}

async function ensureToken() {
  if (!accessToken || Date.now() >= tokenExpiry) {
    await getAccessToken();
  }
}

/* ======================= ФУНКЦИИ ДЛЯ СОЗДАНИЯ РАЗДЕЛОВ И МАТЕРИАЛОВ (ОБРАТНАЯ СВЯЗЬ) ======================= */
async function createSection(courseId, sectionName, iconUrl) {
  await ensureToken();
  const url = `${API_BASE}/courses/${courseId}/sections`;
  const headers = {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type': 'application/json'
  };
  const payload = {
    section: {
      name: sectionName,
      icon_remote_url: iconUrl
    }
  };
  try {
    const response = await axios.post(url, payload, { headers });
    const sectionData = response.data;
    addLogFeedback(`Создан раздел "${sectionName}" для курса ID ${courseId}.`);
    return sectionData.id;
  } catch (error) {
    addLogFeedback(`Ошибка при создании раздела "${sectionName}" для курса ID ${courseId}: ${error}`);
    return null;
  }
}

async function addMaterialWithHyperlink(sectionId, materialName, link) {
  await ensureToken();
  const url = `${API_BASE}/sections/${sectionId}/materials`;
  const headers = {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type': 'application/json'
  };
  const content = {
    blocks: [
      {
        type: "paragraph",
        id: "unique-id",
        data: {
          text: (
            "Уважаемые студенты!<br><br>" +
            "Просим вас принять участие в опросе по нашей дисциплине.<br>" +
            "Ваше мнение очень важно для улучшения качества обучения и организации учебного процесса.<br><br>" +
            `Ссылка на обратную связь по дисциплине: <a href="${link}" target="_blank">ссылка</a><br><br>` +
            "Заранее благодарим за вашу активность!"
          )
        }
      }
    ],
    version: "2.25.0",
    time: Date.now()
  };
  const payload = {
    material: {
      name: materialName,
      description: "Описание отсутствует",
      content: content
    }
  };
  try {
    const response = await axios.post(url, payload, { headers });
    addLogFeedback(`Добавлен материал "${materialName}" в раздел ID ${sectionId}.`);
    return true;
  } catch (error) {
    addLogFeedback(`Ошибка при добавлении материала "${materialName}" в раздел ID ${sectionId}: ${error}`);
    return false;
  }
}

/* ======================= ФУНКЦИОНАЛ "КУРСЫ И ПОТОКИ" ======================= */
async function fetchAllCourses() {
  let courses = [];
  let page = 1;
  const per_page = 100;
  let moreData = true;
  while (moreData) {
    await ensureToken();
    setProgress(`Загрузка курсов: страница ${page}...`);
    addLog(`Запрос курсов, страница ${page}`);
    try {
      const response = await axios.get(COURSES_URL, {
        params: { page, per_page },
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      let data = Array.isArray(response.data)
        ? response.data
        : (response.data.data || []);
      addLog(`Страница ${page}: найдено ${data.length} курсов.`);
      if (Array.isArray(data) && data.length > 0) {
        courses = courses.concat(data);
        page++;
        await delay(500);
      } else {
        addLog('Больше курсов не найдено.');
        moreData = false;
      }
    } catch (error) {
      addLog(`Ошибка получения курсов на странице ${page}: ${error}`);
      moreData = false;
    }
  }
  setProgress('Загрузка курсов завершена.');
  return courses;
}

async function fetchCourseSessions(courseId) {
  const url = `${API_BASE}/courses/${courseId}/course_sessions`;
  let sessions = [];
  let page = 1;
  const per_page = 100;
  let moreData = true;
  while (moreData) {
    await ensureToken();
    try {
      const response = await axios.get(url, {
        params: { page, per_page },
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      let data = Array.isArray(response.data)
        ? response.data
        : (response.data.data || []);
      addLog(`Курс ${courseId} - страница ${page}: найдено ${data.length} сеансов.`);
      if (Array.isArray(data) && data.length > 0) {
        sessions = sessions.concat(data);
        page++;
        await delay(200);
      } else {
        moreData = false;
      }
    } catch (error) {
      addLog(`Ошибка получения сеансов для курса ${courseId} на странице ${page}: ${error}`);
      moreData = false;
    }
  }
  return sessions;
}

async function fetchSessionDetails(courseId, sessionId) {
  const url = `${API_BASE}/courses/${courseId}/course_sessions/${sessionId}`;
  await ensureToken();
  try {
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    let sessionData = response.data;
    if (!sessionData || Array.isArray(sessionData)) {
      sessionData = {};
    }
    const participants = sessionData.participants || [];
    const speakers = participants.filter(p => p.role_name && p.role_name.toLowerCase() === 'докладчик');
    const listenersCount = participants.filter(p => p.role_name && p.role_name.toLowerCase() === 'слушатель').length;
    const courseInfo = sessionData.course || {};
    const ownerName = courseInfo.owner_name || 'Неизвестно';
    let authorsNames = 'Неизвестно';
    if (courseInfo.authors && Array.isArray(courseInfo.authors)) {
      authorsNames = courseInfo.authors.map(a => `${a.last_name || ''} ${a.name || ''}`.trim()).join(', ');
    }
    return { speakers, listenersCount, ownerName, authorsNames };
  } catch (error) {
    addLog(`Ошибка получения деталей сеанса (Course ID: ${courseId}, Session ID: ${sessionId}): ${error}`);
    return { speakers: [], listenersCount: 0, ownerName: 'Неизвестно', authorsNames: 'Неизвестно' };
  }
}

async function processCourseSessions(course) {
  const courseId = course.id;
  let records = [];
  const sessions = await fetchCourseSessions(courseId);
  addLog(`Для курса ${courseId} получено сеансов: ${sessions.length}`);
  if (sessions && sessions.length > 0) {
    for (const session of sessions) {
      const sessionId = session.id;
      const sessionName = session.name || 'Неизвестно';
      const details = await fetchSessionDetails(courseId, sessionId);
      if (details.speakers && details.speakers.length > 0) {
        for (const speaker of details.speakers) {
          records.push({
            course_id: course.id,
            course_name: course.name,
            session_id: sessionId,
            session_name: sessionName,
            user_id: speaker.id,
            fullname: speaker.fullname || `${speaker.last_name || ''} ${speaker.name || ''}`.trim(),
            listeners_count: details.listenersCount,
            owner_name: details.ownerName,
            authors_names: details.authorsNames,
            created_at: course.created_at,
            'Категория': (course.types && Array.isArray(course.types))
              ? course.types.map(t => t.name || 'Неизвестно').join(', ')
              : 'Неизвестно'
          });
        }
      } else {
        records.push({
          course_id: course.id,
          course_name: course.name,
          session_id: sessionId,
          session_name: sessionName,
          user_id: null,
          fullname: null,
          listeners_count: details.listenersCount,
          owner_name: details.ownerName,
          authors_names: details.authorsNames,
          created_at: course.created_at,
          'Категория': (course.types && Array.isArray(course.types))
            ? course.types.map(t => t.name || 'Неизвестно').join(', ')
            : 'Неизвестно'
        });
      }
    }
  }
  return records;
}

function addFullExportRow(data, rowNumber) {
  const tr = document.createElement('tr');
  const tdNum = document.createElement('td');
  tdNum.textContent = rowNumber;
  tr.appendChild(tdNum);
  const fields = ['course_id', 'course_name', 'session_id', 'session_name', 'user_id', 'fullname', 'listeners_count', 'owner_name', 'authors_names', 'created_at', 'Категория'];
  fields.forEach(field => {
    const td = document.createElement('td');
    if (field === 'created_at' && data[field] && new Date(data[field]).toString() !== 'Invalid Date') {
      td.textContent = new Date(data[field]).toLocaleString();
    } else {
      td.textContent = data[field] !== null ? data[field] : '';
    }
    tr.appendChild(td);
  });
  fullExportTableBody.appendChild(tr);
}

async function exportToExcel(data, defaultFileName) {
  addLog("Начало экспорта в Excel");
  const exportData = data.map(item => ({ ...item }));
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
  let filePath = await window.electronAPI.saveFileDialog(defaultFileName);
  if (filePath) {
    XLSX.writeFile(workbook, filePath);
    addLog(`Файл успешно сохранён: ${filePath}`);
  } else {
    addLog('Сохранение файла отменено.');
  }
}

async function fullExport() {
  addLog('Начало полной выгрузки: курсы, сеансы и потоки.');
  setProgress('Получение курсов...');
  const startTime = Date.now();
  let courses = [];
  try {
    courses = await fetchAllCourses();
    addLog(`Всего курсов получено: ${courses.length}`);
  } catch (error) {
    addLog('Ошибка при получении курсов: ' + error);
    return;
  }
  // Применяем фильтры: по дате, типам и по course_id (если загружен файл)
  if (window.filterOptions.startDate || window.filterOptions.endDate || window.filterOptions.courseTypes.length > 0 || (window.filterOptions.courseIds && window.filterOptions.courseIds.length > 0)) {
    addLog(`Текущий фильтр courseIds: ${window.filterOptions.courseIds.join(', ')}`);
    courses = courses.filter(course => {
      if (course.created_at) {
        const courseDate = new Date(course.created_at);
        if (window.filterOptions.startDate) {
          const startDate = new Date(window.filterOptions.startDate);
          if (courseDate < startDate) return false;
        }
        if (window.filterOptions.endDate) {
          const endDate = new Date(window.filterOptions.endDate);
          if (courseDate > endDate) return false;
        }
      }
      if (window.filterOptions.courseTypes.length > 0) {
        if (!course.types || !Array.isArray(course.types)) return false;
        const courseTypeNames = course.types.map(t => (t.name || '').toLowerCase());
        const filterMatch = window.filterOptions.courseTypes.some(filterType => courseTypeNames.includes(filterType.toLowerCase().trim()));
        if (!filterMatch) return false;
      }
      if (window.filterOptions.courseIds && window.filterOptions.courseIds.length > 0) {
        if (!window.filterOptions.courseIds.includes(String(course.id).trim())) return false;
      }
      return true;
    });
    addLog(`После фильтрации осталось курсов: ${courses.length}`);
  }
  fullExportTableBody.innerHTML = '';
  window.fullExportData = [];
  let rowCounter = 0;
  const totalCourses = courses.length;
  for (const [i, course] of courses.entries()) {
    const elapsedSec = Math.round((Date.now() - startTime) / 1000);
    const percent = totalCourses > 0 ? Math.round(((i + 1) / totalCourses) * 100) : 0;
    setProgress(`Обрабатывается курс ${i + 1} из ${totalCourses} (${percent}% завершено). Прошло: ${formatTime(elapsedSec)}.`);
    
    const sessionRecords = await processCourseSessions(course);
    if (sessionRecords && sessionRecords.length > 0) {
      for (const rec of sessionRecords) {
        rowCounter++;
        window.fullExportData.push(rec);
        addFullExportRow(rec, rowCounter);
      }
    } else {
      rowCounter++;
      const record = {
        course_id: course.id,
        course_name: course.name,
        session_id: null,
        session_name: null,
        user_id: null,
        fullname: null,
        listeners_count: 0,
        owner_name: course.owner_name || 'Неизвестно',
        authors_names: (course.authors && Array.isArray(course.authors))
                         ? course.authors.map(a => `${(a.last_name || '').trim()} ${(a.name || '').trim()}`.trim()).join(', ')
                         : 'Неизвестно',
        created_at: course.created_at,
        'Категория': (course.types && Array.isArray(course.types))
                      ? course.types.map(t => t.name || 'Неизвестно').join(', ')
                      : 'Неизвестно'
      };
      window.fullExportData.push(record);
      addFullExportRow(record, rowCounter);
      addLog(`Курс ID ${course.id} ('${course.name}') не имеет сеансов.`);
    }
    await delay(200);
  }
  setProgress('Выгрузка завершена.');
  addLog(`Полная выгрузка завершена. Всего записей: ${window.fullExportData.length}. Общее время: ${formatTime(Math.round((Date.now() - startTime)/1000))}.`);
  document.getElementById('exportFullBtn').disabled = false;
}

/* ======================= ФУНКЦИОНАЛ "РАЗДЕЛЫ И МАТЕРИАЛЫ" ======================= */
async function fetchAllCoursesMaterials() {
  let courses = [];
  let page = 1;
  const per_page = 100;
  let moreData = true;
  while (moreData) {
    await ensureToken();
    setProgressMaterials(`Получение курсов (материалы): страница ${page}...`);
    addLogMaterials(`Запрос курсов для материалов, страница ${page}`);
    try {
      const response = await axios.get(COURSES_URL, {
        params: { page, per_page },
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      let data = Array.isArray(response.data)
        ? response.data
        : (response.data.data || []);
      addLogMaterials(`Страница ${page}: найдено ${data.length} курсов.`);
      if (Array.isArray(data) && data.length > 0) {
        courses = courses.concat(data);
        page++;
        await delay(500);
      } else {
        addLogMaterials('Больше курсов не найдено.');
        moreData = false;
      }
    } catch (error) {
      addLogMaterials(`Ошибка получения курсов на странице ${page}: ${error}`);
      moreData = false;
    }
  }
  setProgressMaterials('Загрузка курсов для материалов завершена.');
  return courses;
}

async function getSections(courseId) {
  await ensureToken();
  try {
    const response = await axios.get(`${API_BASE}/courses/${courseId}/sections`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    return response.data.map(section => ({
      section_id: section.id,
      section_name: section.name || ""
    }));
  } catch (error) {
    addLogMaterials(`Ошибка при получении секций курса ${courseId}: ${error}`);
    return [];
  }
}

async function getMaterials(sectionId) {
  await ensureToken();
  try {
    const response = await axios.get(`${API_BASE}/sections/${sectionId}/materials`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    return response.data.map(material => ({
      material_name: material.name || "",
      file_name: material.file_name || "",
      category: material.category || "",
      material_created_at: material.created_at || "",
      scorm: false
    }));
  } catch (error) {
    addLogMaterials(`Ошибка при получении материалов секции ${sectionId}: ${error}`);
    return [];
  }
}

async function getScorms(sectionId) {
  await ensureToken();
  try {
    const response = await axios.get(`${API_BASE}/sections/${sectionId}/scorms`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    return response.data.map(scorm => ({
      material_name: scorm.name || "",
      file_name: scorm.resource_url || "",
      category: "scorm",
      material_created_at: "",
      scorm: true
    }));
  } catch (error) {
    addLogMaterials(`Ошибка при получении SCORM материалов секции ${sectionId}: ${error}`);
    return [];
  }
}

async function fullExportMaterials() {
  addLogMaterials('Начало выгрузки разделов и материалов.');
  setProgressMaterials('Получение курсов для разделов и материалов...');
  const startTime = Date.now();
  let courses = [];
  try {
    courses = await fetchAllCoursesMaterials();
    addLogMaterials(`Всего курсов получено: ${courses.length}`);
  } catch (error) {
    addLogMaterials('Ошибка при получении курсов: ' + error);
    return;
  }
  // Применяем фильтры: по дате, типам и по course_id (если загружен файл)
  if (window.filterOptionsMaterials.startDate || window.filterOptionsMaterials.endDate || window.filterOptionsMaterials.courseTypes.length > 0 || (window.filterOptionsMaterials.courseIds && window.filterOptionsMaterials.courseIds.length > 0)) {
    addLogMaterials(`Текущий фильтр courseIds: ${window.filterOptionsMaterials.courseIds.join(', ')}`);
    courses = courses.filter(course => {
      if (course.created_at) {
        const courseDate = new Date(course.created_at);
        if (window.filterOptionsMaterials.startDate) {
          const startDate = new Date(window.filterOptionsMaterials.startDate);
          if (courseDate < startDate) return false;
        }
        if (window.filterOptionsMaterials.endDate) {
          const endDate = new Date(window.filterOptionsMaterials.endDate);
          if (courseDate > endDate) return false;
        }
      }
      if (window.filterOptionsMaterials.courseTypes.length > 0) {
        if (!course.types || !Array.isArray(course.types)) return false;
        const courseTypeNames = course.types.map(t => (t.name || '').toLowerCase());
        const filterMatch = window.filterOptionsMaterials.courseTypes.some(filterType => courseTypeNames.includes(filterType.toLowerCase().trim()));
        if (!filterMatch) return false;
      }
      if (window.filterOptionsMaterials.courseIds && window.filterOptionsMaterials.courseIds.length > 0) {
        if (!window.filterOptionsMaterials.courseIds.includes(String(course.id).trim())) return false;
      }
      return true;
    });
    addLogMaterials(`После фильтрации осталось курсов: ${courses.length}`);
  }
  materialsDetailsTableBody.innerHTML = '';
  window.materialsDataDetails = [];
  let detailRowCounter = 0;
  const totalCourses = courses.length;
  for (const [i, course] of courses.entries()) {
    const elapsedSec = Math.round((Date.now() - startTime) / 1000);
    const percent = totalCourses > 0 ? Math.round(((i + 1) / totalCourses) * 100) : 0;
    setProgressMaterials(`Обрабатывается курс ${i + 1} из ${totalCourses} (${percent}% завершено). Прошло: ${formatTime(elapsedSec)}.`);
    
    const courseId = course.id;
    const courseName = course.name;
    const courseCreatedAt = course.created_at || "";
    addLogMaterials(`Обработка курса ID ${courseId} ('${courseName}')...`);
    const sections = await getSections(courseId);
    for (const section of sections) {
      const sectionId = section.section_id;
      const sectionName = section.section_name;
      const materials = await getMaterials(sectionId);
      const scorms = await getScorms(sectionId);
      addLogMaterials(`Секция ID ${sectionId} ('${sectionName}'): найдено ${materials.length} материалов, ${scorms.length} SCORM материалов.`);
      const combined = [...materials, ...scorms];
      if (combined.length > 0) {
        for (const material of combined) {
          detailRowCounter++;
          const record = {
            course_id: courseId,
            course_name: courseName,
            course_created_at: courseCreatedAt,
            section_name: sectionName,
            material_name: material.material_name,
            file_name: material.file_name,
            category: material.category,
            material_created_at: material.material_created_at,
            scorm: material.scorm
          };
          window.materialsDataDetails.push(record);
          const tr = document.createElement('tr');
          const tdNum = document.createElement('td');
          tdNum.textContent = detailRowCounter;
          tr.appendChild(tdNum);
          const fields = ['course_id', 'course_name', 'course_created_at', 'section_name', 'material_name', 'file_name', 'category', 'material_created_at', 'scorm'];
          fields.forEach(field => {
            const td = document.createElement('td');
            td.textContent = record[field] !== null ? record[field] : '';
            tr.appendChild(td);
          });
          materialsDetailsTableBody.appendChild(tr);
        }
      } else {
        detailRowCounter++;
        const record = {
          course_id: courseId,
          course_name: courseName,
          course_created_at: courseCreatedAt,
          section_name: sectionName,
          material_name: "",
          file_name: "",
          category: "",
          material_created_at: "",
          scorm: false
        };
        window.materialsDataDetails.push(record);
        const tr = document.createElement('tr');
        const tdNum = document.createElement('td');
        tdNum.textContent = detailRowCounter;
        tr.appendChild(tdNum);
        const fields = ['course_id', 'course_name', 'course_created_at', 'section_name', 'material_name', 'file_name', 'category', 'material_created_at', 'scorm'];
        fields.forEach(field => {
          const td = document.createElement('td');
          td.textContent = record[field] !== null ? record[field] : '';
          tr.appendChild(td);
        });
        materialsDetailsTableBody.appendChild(tr);
      }
      await delay(100);
    }
    await delay(200);
  }
  setProgressMaterials('Выгрузка завершена.');
  addLogMaterials(`Выгрузка завершена. Всего деталей: ${window.materialsDataDetails.length}. Общее время: ${formatTime(Math.round((Date.now() - startTime)/1000))}.`);
  document.getElementById('exportMaterialsBtn').disabled = false;
}

/* ======================= ФУНКЦИОНАЛ "ОБРАТНАЯ СВЯЗЬ" ======================= */
function createPrefilledUrl(baseUrl, courseName) {
  const params = { answer_long_text_64154736: courseName };
  const encodedParams = Object.keys(params)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');
  return `${baseUrl}?${encodedParams}`;
}

function processFeedbackFile(file) {
  setProgressFeedback('Чтение файла...');
  addLogFeedback('Начало чтения файла обратной связи...');
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    addLogFeedback(`Файл прочитан. Найдено строк: ${jsonData.length}`);
    jsonData.forEach(row => {
      const courseName = row["Название курса в ЛМС"] || "";
      row["Ссылка"] = createPrefilledUrl(FORM_URL, courseName);
    });
    window.feedbackData = jsonData;
    addLogFeedback('Ссылки сгенерированы.');
    setProgressFeedback('Генерация ссылок завершена.');
    setTimeout(() => {
      addLogFeedback('Запуск процесса создания разделов и материалов...');
      awaitCreateFeedback();
    }, 100);
  };
  reader.onerror = function(err) {
    addLogFeedback(`Ошибка чтения файла: ${err}`);
    setProgressFeedback('Ошибка чтения файла.');
  };
  reader.readAsArrayBuffer(file);
}

async function awaitCreateFeedback() {
  window.feedbackResults = [];
  const feedbackSectionIconUrl = 'https://i.ibb.co/dmK60K4/feedback.png';
  addLogFeedback("Начало создания разделов и материалов для обратной связи...");
  const startTime = Date.now();
  const totalRecords = window.feedbackData.length;
  for (const [index, row] of window.feedbackData.entries()) {
    addLogFeedback(`Обработка записи ${index + 1} из ${totalRecords}...`);
    const courseId = row['course_id'];
    const courseName = row["Название курса в ЛМС"] || row["course_name"] || "";
    const link = row["Ссылка"] || "";
    if (!courseId || !courseName || !link) {
      addLogFeedback(`Пропуск записи ${index + 1}: недостаточно данных (course_id: ${courseId}, courseName: ${courseName})`);
      continue;
    }
    const sectionId = await createSection(courseId, "Обратная связь", feedbackSectionIconUrl);
    if (!sectionId) {
      addLogFeedback(`Запись ${index + 1}: Не удалось создать раздел для курса ID ${courseId}.`);
      continue;
    }
    const success = await addMaterialWithHyperlink(sectionId, "Обратная связь по дисциплине", link);
    addLogFeedback(`Запись ${index + 1}: Раздел ID ${sectionId} создан. Материал ${success ? "добавлен" : "не добавлен"}.`);
    window.feedbackResults.push({
      course_id: courseId,
      course_name: courseName,
      section_id: sectionId,
      material_added: success,
      link: link
    });
    await delay(500);
    const elapsedSec = Math.round((Date.now() - startTime) / 1000);
    setProgressFeedback(`Обработано ${index + 1} из ${totalRecords}. Прошло: ${formatTime(elapsedSec)}.`);
  }
  addLogFeedback("Создание разделов и материалов 'Обратная связь' завершено.");
  exportFeedbackBtn.disabled = false;
}

async function exportFeedbackResultsToExcel(defaultFileName) {
  addLogFeedback("Начало экспорта результата в Excel");
  const worksheet = XLSX.utils.json_to_sheet(window.feedbackResults);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Results");
  let filePath = await window.electronAPI.saveFileDialog(defaultFileName);
  if (filePath) {
    XLSX.writeFile(workbook, filePath);
    addLogFeedback(`Файл успешно сохранён: ${filePath}`);
  } else {
    addLogFeedback('Сохранение файла отменено.');
  }
}

/* ======================= ОБРАБОТЧИКИ КНОПОК ======================= */
// "Курсы и потоки"
document.getElementById('fullExportBtn').addEventListener('click', async () => {
  document.getElementById('exportFullBtn').disabled = true;
  await fullExport();
});
document.getElementById('exportFullBtn').addEventListener('click', async () => {
  if (window.fullExportData && window.fullExportData.length > 0) {
    await exportToExcel(window.fullExportData, `course_sessions_speakers_${new Date().toISOString().slice(0,10)}.xlsx`);
  } else {
    addLog('Нет данных для экспорта.');
  }
});

// "Разделы и материалы"
document.getElementById('fullExportMaterialsBtn').addEventListener('click', async () => {
  document.getElementById('exportMaterialsBtn').disabled = true;
  await fullExportMaterials();
});
document.getElementById('exportMaterialsBtn').addEventListener('click', async () => {
  if (window.materialsDataDetails && window.materialsDataDetails.length > 0) {
    await exportToExcel(window.materialsDataDetails, `courses_materials_${new Date().toISOString().slice(0,10)}.xlsx`);
  } else {
    addLogMaterials('Нет данных для экспорта.');
  }
});

// "Обратная связь"
processFeedbackBtn.addEventListener('click', () => {
  const files = feedbackFileInput.files;
  if (!files || files.length === 0) {
    addLogFeedback('Выберите файл для обработки.');
    return;
  }
  exportFeedbackBtn.disabled = true;
  processFeedbackFile(files[0]);
});

exportFeedbackBtn.addEventListener('click', async () => {
  if (window.feedbackResults && window.feedbackResults.length > 0) {
    await exportFeedbackResultsToExcel(`feedback_results_${new Date().toISOString().slice(0,10)}.xlsx`);
  } else {
    addLogFeedback('Нет данных для экспорта.');
  }
});

/* ======================= ФУНКЦИОНАЛ ФИЛЬТРОВ ======================= */
async function fetchCourseTypes() {
  try {
    await ensureToken();
    const response = await axios.get(COURSE_TYPES_URL, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const types = response.data;
    const container = document.getElementById('courseTypesContainer');
    container.innerHTML = '';
    types.forEach(type => {
      const label = document.createElement('label');
      label.style.display = 'block';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = type.name;
      checkbox.name = 'courseTypeCheckbox';
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(` ${type.name}`));
      container.appendChild(label);
    });
    addLog("Типы курсов успешно загружены для фильтрации.");
  } catch (error) {
    addLog(`Ошибка получения типов курсов: ${error}`);
  }
}

document.getElementById('filterBtn').addEventListener('click', async () => {
  currentFilterTarget = 'courses';
  await fetchCourseTypes();
  if (courseIdFileInput.files && courseIdFileInput.files.length > 0) {
    processCourseIdFilterFile(courseIdFileInput.files[0]);
  } else {
    window.filterOptions.courseIds = [];
  }
  document.getElementById('filterModal').style.display = 'block';
});

document.getElementById('filterMaterialsBtn').addEventListener('click', async () => {
  currentFilterTarget = 'materials';
  await fetchCourseTypes();
  if (courseIdFileInput.files && courseIdFileInput.files.length > 0) {
    processCourseIdFilterFile(courseIdFileInput.files[0]);
  } else {
    window.filterOptionsMaterials.courseIds = [];
  }
  document.getElementById('filterModal').style.display = 'block';
});

document.getElementById('applyFiltersBtn').addEventListener('click', () => {
  const startVal = document.getElementById('startDate').value || null;
  const endVal = document.getElementById('endDate').value || null;
  const checkboxes = document.querySelectorAll('input[name="courseTypeCheckbox"]:checked');
  const typesArr = Array.from(checkboxes).map(cb => cb.value);
  if (currentFilterTarget === 'courses') {
    window.filterOptions.startDate = startVal;
    window.filterOptions.endDate = endVal;
    window.filterOptions.courseTypes = typesArr;
    addLog('Фильтры применены для Курсов и потоков.');
  } else {
    window.filterOptionsMaterials.startDate = startVal;
    window.filterOptionsMaterials.endDate = endVal;
    window.filterOptionsMaterials.courseTypes = typesArr;
    addLogMaterials('Фильтры применены для Разделов и материалов.');
  }
  document.getElementById('filterModal').style.display = 'none';
});

document.getElementById('cancelFiltersBtn').addEventListener('click', () => {
  document.getElementById('filterModal').style.display = 'none';
});
window.onclick = function(event) {
  if (event.target == document.getElementById('filterModal')) {
    document.getElementById('filterModal').style.display = 'none';
  }
};

// NEW: Функция для обработки файла с course_id в фильтрах (переименована)
function processCourseIdFilterFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    const ids = jsonData
      .map(row => String(row.course_id).trim())
      .filter(id => id.length > 0);
    if (ids.length > 0) {
      addLog(`Фильтр по course_id: найдено ${ids.length} записей.`);
      window.filterOptions.courseIds = ids;
      window.filterOptionsMaterials.courseIds = ids;
    } else {
      addLog("Файл с course_id не содержит данных.");
    }
  };
  reader.onerror = function(err) {
    addLog(`Ошибка чтения файла с course_id: ${err}`);
  };
  reader.readAsArrayBuffer(file);
}


