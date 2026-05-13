const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { execSync } = require('child_process');

// Load trip config (optional — parser works without it)
let tripConfig = {};
try {
    tripConfig = JSON.parse(fs.readFileSync(path.join(__dirname, 'trip-config.json'), 'utf-8'));
} catch (e) {}

// Top-level dayOfYear helper (also used inside merge step)
const _DAYS_IN_MONTH = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
function _dayOfYear(day, month) {
    let n = day;
    for (let i = 1; i < month; i++) n += _DAYS_IN_MONTH[i];
    return n;
}

// Поддержка нескольких URL через аргументы командной строки
// Использование: node aviasales-parser.js URL1 URL2 URL3 ...
// Или: node aviasales-parser.js (использует URL по умолчанию)
const defaultUrl = 'https://www.aviasales.ru/search/MOW2102DXB2502MRU0503MOW2';
const urls = process.argv.slice(2);
if (urls.length === 0) {
    urls.push(defaultUrl);
}

// Функция для извлечения дат и маршрута из URL
function generateFileName(url) {
    // Парсим URL: 3-сегментный MOW2102DXB2502MRU0503MOW2, 2-сегментный MRU0503IST0503MOW2, или 1-сегментный AER1206KJA2
    const match3 = url.match(/([A-Z]{3})(\d{4})([A-Z]{3})(\d{4})([A-Z]{3})(\d{4})([A-Z]{3})/);
    const match2 = !match3 && url.match(/\/([A-Z]{3})(\d{4})([A-Z]{3})(\d{4})([A-Z]{3})\d*$/);
    const match1 = !match3 && !match2 && url.match(/\/([A-Z]{3})(\d{4})([A-Z]{3})(\d+)$/);

    let dates = 'unknown';
    let route = 'unknown';
    let startDate = 'unknown';

    let cities = ['MOW', 'DXB', 'MRU', 'MOW']; // По умолчанию
    if (match3) {
        const [, city1, date1, city2, date2, city3, date3, city4] = match3;
        dates = `${date1}-${date2}-${date3}`;
        route = `${city1}-${city2}-${city3}`;
        startDate = date1;
        cities = [city1, city2, city3, city4];
    } else if (match2) {
        const [, city1, date1, city2, date2, city3] = match2;
        dates = `${date1}-${date2}`;
        route = `${city1}-${city2}-${city3}`;
        startDate = date1;
        cities = [city1, city2, city3];
    } else if (match1) {
        const [, city1, date1, city2] = match1;
        dates = date1;
        route = `${city1}-${city2}`;
        startDate = date1;
        cities = [city1, city2];
    }

    const now = new Date();
    const dateStr = (now.getMonth() + 1).toString().padStart(2, '0') + '-' + now.getDate().toString().padStart(2, '0');
    const timeStr = now.toTimeString().slice(0, 5).replace(':', '-');
    const timestamp = `${dateStr}_${timeStr}`; // MM-DD_HH-MM

    return {
        excel: `output/as_${dates}_${route}_${timestamp}.xlsx`,
        json: `output/as_${dates}_${route}_${timestamp}.json`,
        report: `output/as_${dates}_${route}_${timestamp}.txt`,
        screenshot: `output/as_${dates}_${route}_${timestamp}.png`,
        startDate: startDate,
        dates: dates,
        route: route,
        cities: cities // [city1, city2, city3, city4] для маршрутов в Excel
    };
}

// Функция парсинга одной страницы
async function parseOnePage(page, searchUrl, isFirstUrl) {
    const fileNames = generateFileName(searchUrl);

    console.log('\n' + '═'.repeat(70));
    console.log(`📍 ПАРСИНГ: ${fileNames.dates} | ${fileNames.route}`);
    console.log('═'.repeat(70));
    console.log(`Файл: ${fileNames.excel}\n`);

    await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });

    // Капча только для первого URL
    if (isFirstUrl) {
        console.log('┌' + '─'.repeat(68) + '┐');
        console.log('│  ⏸️  РЕШИТЕ CAPTCHA В ОТКРЫВШЕМСЯ БРАУЗЕРЕ!                          │');
        console.log('│                                                                    │');
        console.log('│  У вас есть 60 секунд на решение captcha и загрузку билетов       │');
        console.log('└' + '─'.repeat(68) + '┘\n');

        const waitTime = 60;
        for (let i = 0; i < waitTime; i += 15) {
            console.log(`   ⏳ Осталось ~${waitTime - i} секунд...`);
            await page.waitForTimeout(15000);
        }
        console.log('\n✓ Время истекло! Проверяем страницу...');
    } else {
        // Для следующих URL — короткое ожидание
        console.log('⏳ Ожидание загрузки (8 сек)...');
        await page.waitForTimeout(8000);
    }

    // Быстрая проверка загрузки (макс 5 сек)
    console.log('Проверяем загрузку...');
    for (let i = 0; i < 5; i++) {
        const priceCount = await page.evaluate(() => {
            return (document.body.textContent.match(/\d{2,6}\s*₽/g) || []).length;
        });
        if (priceCount > 50) {
            console.log(`✓ Найдено ${priceCount} цен, продолжаем`);
            break;
        }
        await page.waitForTimeout(1000);
    }

    // Нажимаем "Показать ещё" (быстро)
    console.log('Загружаем все билеты...');
    let showMoreClicks = 0;

    while (showMoreClicks < 20) {
        const buttonFound = await page.evaluate(() => {
            const buttons = Array.from(document.querySelectorAll('button, [role="button"]'));
            const showMoreButton = buttons.find(btn =>
                btn.textContent.includes('Показать ещё') ||
                btn.textContent.includes('Загрузить ещё') ||
                btn.textContent.includes('Ещё варианты')
            );
            if (showMoreButton) {
                showMoreButton.click();
                return true;
            }
            return false;
        });

        if (!buttonFound) break;
        showMoreClicks++;
        if (showMoreClicks % 5 === 0) console.log(`  Показать ещё: ${showMoreClicks}...`);
        await page.waitForTimeout(1500); // Было 3000
    }
    if (showMoreClicks > 0) console.log(`✓ Загружено (${showMoreClicks} кликов)`);

    // Быстрая прокрутка
    console.log('Прокрутка страницы...');
    let previousHeight = 0;
    for (let i = 0; i < 10; i++) {
        const currentHeight = await page.evaluate(() => document.body.scrollHeight);
        if (currentHeight === previousHeight && i > 2) break;
        await page.evaluate(() => window.scrollBy(0, 1500));
        await page.waitForTimeout(800); // Было 2000
        previousHeight = currentHeight;
    }
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.waitForTimeout(500);

    console.log('\nИзвлекаем данные о билетах...');

    // Извлечение билетов
    const tickets = await page.evaluate(() => {
        const results = [];
        const allElements = document.querySelectorAll('*');
        const priceElements = Array.from(allElements).filter(el => {
            const text = el.textContent;
            return text && /\d+\s*₽/.test(text) && el.children.length > 5;
        });

        const ticketCards = document.querySelectorAll(
            '[data-test-id*="ticket"], [data-test-id*="card"], [data-test-id*="proposal"]'
        );

        const potentialCards = new Set([...ticketCards, ...priceElements.slice(0, 50)]);

        Array.from(potentialCards).forEach((card, index) => {
            try {
                const ticketData = {
                    index: index + 1,
                    price: null,
                    priceValue: null,
                    segments: [],
                    rawText: ''
                };

                const priceSelectors = ['[data-test-id*="price"]', '[class*="price"]', '[class*="Price"]'];
                for (const selector of priceSelectors) {
                    const priceElement = card.querySelector(selector);
                    if (priceElement && priceElement.textContent.match(/\d/)) {
                        ticketData.price = priceElement.textContent.trim();
                        const match = ticketData.price.match(/(\d[\d\s]*)/);
                        if (match) {
                            ticketData.priceValue = parseInt(match[1].replace(/\s/g, ''));
                        }
                        break;
                    }
                }

                const text = card.textContent;
                ticketData.rawText = text.replace(/\s+/g, ' ').trim();

                const segmentElements = card.querySelectorAll('[class*="segment"], [class*="Segment"], [class*="leg"]');
                if (segmentElements.length > 0) {
                    segmentElements.forEach((segment, segIndex) => {
                        ticketData.segments.push({ segmentNumber: segIndex + 1, isDirect: true });
                    });
                }

                if (ticketData.price || ticketData.segments.length > 0) {
                    results.push(ticketData);
                }
            } catch (error) {}
        });

        return results;
    });

    console.log(`✓ Найдено билетов: ${tickets.length}`);

    // Сохраняем JSON
    const allDataFile = fileNames.json.replace('.json', '-all.json');
    fs.writeFileSync(allDataFile, JSON.stringify(tickets, null, 2), 'utf-8');

    // Функция парсинга деталей
    function parseFlightDetails(rawText) {
        if (!rawText) return null;
        const text = rawText.replace(/\s+/g, ' ').replace(/⁠/g, ' ');
        const times = text.match(/\d{2}:\d{2}/g) || [];
        const dates = text.match(/\d{1,2}\s*(?:янв|фев|мар|апр|мая|июн|июл|авг|сен|окт|ноя|дек)/gi) || [];
        const airports = text.match(/[A-Z]{3}/g) || [];
        const flightInfos = text.match(/(\d+\s*[чд]\s*\d*\s*[чм]?\s*в пути[,\s]*(прямой|прямым|\d+\s*пересадк[аи]?))/gi) || [];

        function getFlightType(info) {
            if (!info) return '?';
            const lower = info.toLowerCase();
            if (lower.includes('прямой') || lower.includes('прямым')) return 'ПРЯМОЙ ✓';
            const stops = lower.match(/(\d+)\s*пересадк/);
            if (stops) return `${stops[1]} пересадка`;
            return '?';
        }

        function getDuration(info) {
            if (!info) return '';
            const match = info.match(/(\d+\s*[чд]\s*\d*\s*[чм]?)/);
            return match ? match[1].trim() : '';
        }

        // Извлекаем транзитные аэропорты для каждого сегмента
        // Паттерн: "N пересадк[аи]" затем коды аэропортов (транзит + конечный)
        function getTransitAirports(segmentIndex) {
            // Ищем паттерн "пересадк" и после него коды аэропортов
            const pattern = /(\d+)\s*пересадк[аи]?\s*([A-Z]{3}(?:[A-Z]{3})*)/gi;
            const matches = [...text.matchAll(pattern)];
            if (matches[segmentIndex]) {
                const [, numStops, codes] = matches[segmentIndex];
                // Извлекаем коды (каждый по 3 символа)
                const allCodes = codes.match(/[A-Z]{3}/g) || [];
                // Все кроме последнего - транзитные
                if (allCodes.length > 1) {
                    return allCodes.slice(0, -1).join(', ');
                } else if (allCodes.length === 1 && parseInt(numStops) > 0) {
                    return allCodes[0];
                }
            }
            return '';
        }

        const c = fileNames.cities;
        return {
            seg1_depart: times[0] || '', seg1_arrive: times[1] || '',
            seg1_date_depart: dates[0] || '', seg1_date_arrive: dates[1] || '',
            seg1_from: airports[0] || c[0] || '', seg1_to: airports[1] || c[1] || '',
            seg1_duration: getDuration(flightInfos[0]), seg1_type: getFlightType(flightInfos[0]),
            seg1_transit: getTransitAirports(0),

            seg2_depart: times[2] || '', seg2_arrive: times[3] || '',
            seg2_date_depart: dates[2] || '', seg2_date_arrive: dates[3] || '',
            seg2_from: airports[2] || c[1] || '', seg2_to: c[2] || '',
            seg2_duration: getDuration(flightInfos[1]), seg2_type: getFlightType(flightInfos[1]),
            seg2_transit: getTransitAirports(1),

            seg3_depart: times[4] || '', seg3_arrive: times[5] || '',
            seg3_date_depart: dates[4] || '', seg3_date_arrive: dates[5] || '',
            seg3_from: airports[4] || c[2] || '', seg3_to: airports[airports.length - 1] || c[3] || c[2] || '',
            seg3_duration: getDuration(flightInfos[2]), seg3_type: getFlightType(flightInfos[2]),
            seg3_transit: getTransitAirports(2)
        };
    }

    // Подготовка данных для Excel
    const allTicketsData = tickets.map((ticket, idx) => {
        const d = parseFlightDetails(ticket.rawText || '');
        // Для маршрутов MRU→* показываем все билеты (пользователь фильтрует в Excel)
        // Для оригинального маршрута MOW→DXB→MRU→MOW применяем критерии
        const cities = fileNames.cities;
        const matchesCriteria = d ? checkCriteria(d, cities) : false;

        // Используем города из URL для маршрутов (вместо извлечения из текста)
        const has2seg = cities.length >= 3;
        const has3seg = cities.length >= 4;

        const row = {
            '✓': matchesCriteria ? '✓' : '',
            '№': idx + 1,
            'Цена': ticket.price || '',
            'Р1': `${cities[0]}→${cities[1]}`,
            'Р1 Дата': d ? d.seg1_date_depart : '',
            'Р1 Вылет': d ? d.seg1_depart : '',
            'Р1 Дата2': d ? d.seg1_date_arrive : '',
            'Р1 Прилёт': d ? d.seg1_arrive : '',
            'Р1 Время': d ? d.seg1_duration : '',
            'Р1 Тип': d ? d.seg1_type : '',
            'Р1 Транзит': d ? d.seg1_transit : ''
        };

        if (has2seg) {
            row['Р2'] = `${cities[1]}→${cities[2]}`;
            row['Р2 Дата'] = d ? d.seg2_date_depart : '';
            row['Р2 Вылет'] = d ? d.seg2_depart : '';
            row['Р2 Дата2'] = d ? d.seg2_date_arrive : '';
            row['Р2 Прилёт'] = d ? d.seg2_arrive : '';
            row['Р2 Время'] = d ? d.seg2_duration : '';
            row['Р2 Тип'] = d ? d.seg2_type : '';
            row['Р2 Транзит'] = d ? d.seg2_transit : '';
        }

        if (has3seg) {
            row['Р3'] = `${cities[2]}→${cities[3]}`;
            row['Р3 Дата'] = d ? d.seg3_date_depart : '';
            row['Р3 Вылет'] = d ? d.seg3_depart : '';
            row['Р3 Дата2'] = d ? d.seg3_date_arrive : '';
            row['Р3 Прилёт'] = d ? d.seg3_arrive : '';
            row['Р3 Время'] = d ? d.seg3_duration : '';
            row['Р3 Тип'] = d ? d.seg3_type : '';
            row['Р3 Транзит'] = d ? d.seg3_transit : '';
        }

        return row;
    });

    // Сортировка
    function timeToMinutes(timeStr) {
        if (!timeStr) return 9999;
        const match = String(timeStr).match(/(\d{1,2}):(\d{2})/);
        if (!match) return 9999;
        return parseInt(match[1]) * 60 + parseInt(match[2]);
    }

    allTicketsData.sort((a, b) => {
        if (a['✓'] === '✓' && b['✓'] !== '✓') return -1;
        if (a['✓'] !== '✓' && b['✓'] === '✓') return 1;
        const priceA = parseInt((a['Цена'] || '999999').replace(/\D/g, '')) || 999999;
        const priceB = parseInt((b['Цена'] || '999999').replace(/\D/g, '')) || 999999;
        if (priceA !== priceB) return priceA - priceB;
        const time1A = timeToMinutes(a['Р1 Вылет']);
        const time1B = timeToMinutes(b['Р1 Вылет']);
        if (time1A !== time1B) return time1A - time1B;
        return 0;
    });

    // Создаём Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(allTicketsData);
    const segCols = [{ wch: 9 }, { wch: 6 }, { wch: 5 }, { wch: 6 }, { wch: 5 }, { wch: 8 }, { wch: 10 }, { wch: 6 }];
    const baseCols = [{ wch: 3 }, { wch: 4 }, { wch: 12 }];
    const numSegs = fileNames.cities.length - 1;
    ws['!cols'] = [...baseCols, ...Array(numSegs).fill(segCols).flat()];
    XLSX.utils.book_append_sheet(wb, ws, 'Билеты');

    // Сводка (только подходящие)
    const byPrice = {};
    allTicketsData.forEach(row => {
        const price = row['Цена'] || 'Без цены';
        const matchesCriteria = row['✓'] === '✓';
        if (!byPrice[price]) {
            byPrice[price] = {
                price, matchesCriteria, count: 0,
                r1_routes: new Set(), r1_dates: new Set(), r1_times: new Set(), r1_transits: new Set(),
                r2_routes: new Set(), r2_dates: new Set(), r2_times: new Set(), r2_transits: new Set(),
                r3_routes: new Set(), r3_dates: new Set(), r3_times: new Set(), r3_transits: new Set()
            };
        }
        byPrice[price].count++;
        if (matchesCriteria) byPrice[price].matchesCriteria = true;
        if (row['Р1']) byPrice[price].r1_routes.add(row['Р1']);
        if (row['Р1 Дата']) byPrice[price].r1_dates.add(row['Р1 Дата']);
        if (row['Р1 Вылет']) byPrice[price].r1_times.add(row['Р1 Вылет']);
        if (row['Р1 Транзит']) byPrice[price].r1_transits.add(row['Р1 Транзит']);
        if (row['Р2']) byPrice[price].r2_routes.add(row['Р2']);
        if (row['Р2 Дата']) byPrice[price].r2_dates.add(row['Р2 Дата']);
        if (row['Р2 Вылет']) byPrice[price].r2_times.add(row['Р2 Вылет']);
        if (row['Р2 Транзит']) byPrice[price].r2_transits.add(row['Р2 Транзит']);
        if (row['Р3']) byPrice[price].r3_routes.add(row['Р3']);
        if (row['Р3 Дата']) byPrice[price].r3_dates.add(row['Р3 Дата']);
        if (row['Р3 Вылет']) byPrice[price].r3_times.add(row['Р3 Вылет']);
        if (row['Р3 Транзит']) byPrice[price].r3_transits.add(row['Р3 Транзит']);
    });

    // Формируем компактную строку пересадок
    function formatTransits(r1, r2, r3) {
        const parts = [];
        const t1 = [...r1].filter(t => t).join('/');
        const t2 = [...r2].filter(t => t).join('/');
        const t3 = [...r3].filter(t => t).join('/');
        if (t1) parts.push(`Р1:${t1}`);
        if (t2) parts.push(`Р2:${t2}`);
        if (t3) parts.push(`Р3:${t3}`);
        return parts.length > 0 ? parts.join(' ') : '—';
    }

    const summaryData = Object.values(byPrice)
        .filter(item => item.matchesCriteria)
        .map(item => ({
            'Цена': item.price,
            'Кол-во': item.count,
            'Пересадка': formatTransits(item.r1_transits, item.r2_transits, item.r3_transits),
            'Р1': [...item.r1_routes].join(', '),
            'Р1 Дата': [...item.r1_dates].sort().join(', '),
            'Р1 Вылет': [...item.r1_times].sort().join(', '),
            'Р2': [...item.r2_routes].join(', '),
            'Р2 Дата': [...item.r2_dates].sort().join(', '),
            'Р2 Вылет': [...item.r2_times].sort().join(', '),
            'Р3': [...item.r3_routes].join(', '),
            'Р3 Дата': [...item.r3_dates].sort().join(', '),
            'Р3 Вылет': [...item.r3_times].sort().join(', ')
        }))
        .sort((a, b) => {
            const priceA = parseInt((a['Цена'] || '999999').replace(/\D/g, '')) || 999999;
            const priceB = parseInt((b['Цена'] || '999999').replace(/\D/g, '')) || 999999;
            return priceA - priceB;
        });

    const summaryWs = XLSX.utils.json_to_sheet(summaryData);
    summaryWs['!cols'] = [
        { wch: 12 },  // Цена
        { wch: 5 },   // Кол-во
        { wch: 12 },  // Пересадка
        { wch: 10 },  // Р1 (маршрут)
        { wch: 8 },   // Р1 Дата
        { wch: 14 },  // Р1 Вылет
        { wch: 10 },  // Р2
        { wch: 8 },   // Р2 Дата
        { wch: 14 },  // Р2 Вылет
        { wch: 10 },  // Р3
        { wch: 8 },   // Р3 Дата
        { wch: 14 }   // Р3 Вылет
    ];
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Сводка');

    // Сохраняем
    XLSX.writeFile(wb, fileNames.excel);
    console.log(`✓ Excel сохранён: ${fileNames.excel}`);
    console.log(`✓ Сводка: ${summaryData.length} подходящих цен`);

    // Скриншот
    await page.screenshot({ path: fileNames.screenshot, fullPage: true });

    return fileNames;
}

// Проверка соответствия билета критериям из trip-config.json
function checkCriteria(d, cities) {
    function segOk(type, criterion) {
        if (!criterion || criterion === 'любой') return true;
        if (criterion === 'прямой' || criterion === 'только прямой') return !!(type && type.includes('ПРЯМОЙ'));
        if (criterion === 'прямой или 1 пересадка') return !!(type && (type.includes('ПРЯМОЙ') || type.includes('1 пересадка')));
        return true;
    }
    const criteria = (tripConfig['критерии'] || tripConfig.criteria) || {};
    const k1 = (cities[0] && cities[1]) ? `${cities[0]}_${cities[1]}` : '';
    const k2 = (cities[1] && cities[2]) ? `${cities[1]}_${cities[2]}` : '';
    const k3 = (cities[2] && cities[3]) ? `${cities[2]}_${cities[3]}` : '';
    return segOk(d.seg1_type, criteria[k1]) &&
           (!k2 || segOk(d.seg2_type, criteria[k2])) &&
           (!k3 || segOk(d.seg3_type, criteria[k3]));
}

// Главная функция
async function parseAviasales() {
    console.log('\n' + '█'.repeat(70));
    console.log('  AVIASALES PARSER — Парсинг нескольких дат');
    console.log('█'.repeat(70));
    console.log(`\n📋 Запланировано: ${urls.length} URL\n`);
    urls.forEach((url, i) => {
        const fn = generateFileName(url);
        console.log(`   ${i + 1}. ${fn.dates} | ${fn.route}`);
    });

    // Закрываем Excel
    console.log('\nЗакрываем Excel если открыт...');
    try {
        execSync('pkill -f "Microsoft Excel" 2>/dev/null || true', { stdio: 'ignore' });
        await new Promise(r => setTimeout(r, 1000));
    } catch (e) {}

    console.log('Запуск браузера...\n');
    const browser = await chromium.launch({
        headless: false,
        args: ['--start-maximized']
    });

    const context = await browser.newContext({
        viewport: null,
        userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
    });

    const page = await context.newPage();
    const createdFiles = [];

    try {
        for (let i = 0; i < urls.length; i++) {
            const isFirstUrl = (i === 0);
            const fileNames = await parseOnePage(page, urls[i], isFirstUrl);
            createdFiles.push(fileNames);

            // Короткая пауза между запросами
            if (i < urls.length - 1) {
                console.log('\n→ Следующая дата...\n');
                await page.waitForTimeout(2000);
            }
        }

        // Итоги
        console.log('\n' + '█'.repeat(70));
        console.log('  ✅ ВСЕ ДАТЫ ОБРАБОТАНЫ!');
        console.log('█'.repeat(70));
        console.log('\nСозданные файлы:');
        createdFiles.forEach((f, i) => {
            console.log(`   ${i + 1}. 📊 ${f.excel}`);
        });

        // Если несколько файлов — объединяем сводки
        if (createdFiles.length > 1) {
            console.log('\n📋 Объединяем сводки...');
            const allData = [];

            createdFiles.forEach(f => {
                const wb = XLSX.readFile(f.excel);
                if (!wb.Sheets['Сводка']) return;

                const data = XLSX.utils.sheet_to_json(wb.Sheets['Сводка']);

                // Формируем лейбл: для MOW→DXB→MRU считаем ночи, для остальных — просто маршрут+даты
                const match3seg = f.dates.match(/(\d{2})(\d{2})-(\d{2})(\d{2})-(\d{2})(\d{2})/);
                let label = `${f.route} | ${f.dates}`;
                if (match3seg && f.route === 'MOW-DXB-MRU') {
                    const daysInMonth = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
                    function dayOfYear(d, m) {
                        let t = d;
                        for (let i = 1; i < m; i++) t += daysInMonth[i];
                        return t;
                    }
                    const d1 = parseInt(match3seg[1]), m1 = parseInt(match3seg[2]);
                    const d2 = parseInt(match3seg[3]), m2 = parseInt(match3seg[4]);
                    const d3 = parseInt(match3seg[5]), m3 = parseInt(match3seg[6]);
                    const dxbN = dayOfYear(d2, m2) - dayOfYear(d1, m1);
                    const mruN = dayOfYear(d3, m3) - dayOfYear(d2, m2);
                    label = `${match3seg[1]}.${match3seg[2]} → ${match3seg[3]}.${match3seg[4]} → ${match3seg[5]}.${match3seg[6]}  |  Дубай ${dxbN}н, Маврикий ${mruN}н`;
                }

                allData.push({ 'Цена': `═══ ${label} ═══` });
                data.forEach(row => allData.push(row));
                allData.push({});
            });

            const summaryWb = XLSX.utils.book_new();
            const summaryWs = XLSX.utils.json_to_sheet(allData);
            summaryWs['!cols'] = [
                { wch: 25 }, { wch: 5 }, { wch: 10 }, { wch: 8 }, { wch: 14 },
                { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 14 }
            ];
            XLSX.utils.book_append_sheet(summaryWb, summaryWs, 'Все сводки');

            // "По датам": сетка дата вылета × кол-во ночей → мин. цена
            try {
                const isOneWay = createdFiles[0].cities.length === 2;
                const dateGrid = {};
                const nightsSet = new Set();
                createdFiles.forEach(f => {
                    const m1 = f.dates.match(/^(\d{2})(\d{2})$/);
                    const m2 = f.dates.match(/^(\d{2})(\d{2})-(\d{2})(\d{2})$/);
                    const m3 = f.dates.match(/^(\d{2})(\d{2})-(\d{2})(\d{2})-(\d{2})(\d{2})$/);
                    let depDay, depMonth, retDay, retMonth;
                    if (m1)      { depDay=+m1[1]; depMonth=+m1[2]; retDay=depDay; retMonth=depMonth; }
                    else if (m2) { depDay=+m2[1]; depMonth=+m2[2]; retDay=+m2[3]; retMonth=+m2[4]; }
                    else if (m3) { depDay=+m3[1]; depMonth=+m3[2]; retDay=+m3[5]; retMonth=+m3[6]; }
                    else return;

                    const nights = isOneWay ? 0 : _dayOfYear(retDay, retMonth) - _dayOfYear(depDay, depMonth);
                    const depKey = `${String(depDay).padStart(2,'0')}.${String(depMonth).padStart(2,'0')}`;
                    nightsSet.add(nights);
                    if (!dateGrid[depKey]) dateGrid[depKey] = {};

                    const wb2 = XLSX.readFile(f.excel);
                    const sheet = wb2.Sheets['Сводка'];
                    if (!sheet) return;
                    const rows = XLSX.utils.sheet_to_json(sheet);
                    // Only real price rows: must contain ₽ and be > 5000 (filters out date strings like "30 апр — 2 мая")
                    const prices = rows
                        .filter(r => (r['Цена']||'').includes('₽'))
                        .map(r => parseInt((r['Цена']||'').replace(/\D/g,'')))
                        .filter(p => p > 5000 && p < 300000);
                    if (prices.length > 0) {
                        const minP = Math.min(...prices);
                        if (!dateGrid[depKey][nights] || minP < dateGrid[depKey][nights]) {
                            dateGrid[depKey][nights] = minP;
                        }
                    }
                });

                const sortedNights = [...nightsSet].sort((a, b) => a - b);
                const sortedDates  = Object.keys(dateGrid).sort((a, b) => {
                    const [da, ma] = a.split('.').map(Number);
                    const [db, mb] = b.split('.').map(Number);
                    return _dayOfYear(da, ma) - _dayOfYear(db, mb);
                });

                const gridRows = sortedDates.map(date => {
                    const row = { 'Дата вылета': date };
                    if (isOneWay) {
                        const p = dateGrid[date][0];
                        row['Мин. цена'] = p ? p.toLocaleString('ru-RU') + ' ₽' : '—';
                    } else {
                        sortedNights.forEach(n => {
                            const p = dateGrid[date][n];
                            row[`${n} ноч.`] = p ? p.toLocaleString('ru-RU') + ' ₽' : '—';
                        });
                    }
                    return row;
                });

                const gridWs = XLSX.utils.json_to_sheet(gridRows);
                gridWs['!cols'] = [{ wch: 14 }, ...(isOneWay ? [{ wch: 16 }] : sortedNights.map(() => ({ wch: 16 })))];
                XLSX.utils.book_append_sheet(summaryWb, gridWs, 'По датам');
                console.log(`✓ Сводка "По датам": ${sortedDates.length} дат${isOneWay ? '' : ` × ${sortedNights.length} вар. ночей`}`);

                // "Все варианты": полный список всех рейсов отсортированный по цене с временем вылета
                const allOptions = [];
                createdFiles.forEach(f2 => {
                    const m1b = f2.dates.match(/^(\d{2})(\d{2})$/);
                    const m2b = f2.dates.match(/^(\d{2})(\d{2})-(\d{2})(\d{2})$/);
                    const m3b = f2.dates.match(/^(\d{2})(\d{2})-(\d{2})(\d{2})-(\d{2})(\d{2})$/);
                    let dD, dM, rD, rM;
                    if (m1b)      { dD=+m1b[1]; dM=+m1b[2]; rD=dD; rM=dM; }
                    else if (m2b) { dD=+m2b[1]; dM=+m2b[2]; rD=+m2b[3]; rM=+m2b[4]; }
                    else if (m3b) { dD=+m3b[1]; dM=+m3b[2]; rD=+m3b[5]; rM=+m3b[6]; }
                    else return;
                    const n2 = isOneWay ? 0 : _dayOfYear(rD, rM) - _dayOfYear(dD, dM);
                    const dk = `${String(dD).padStart(2,'0')}.${String(dM).padStart(2,'0')}`;

                    const wb3 = XLSX.readFile(f2.excel);
                    const s2 = wb3.Sheets['Сводка'];
                    if (!s2) return;
                    XLSX.utils.sheet_to_json(s2).forEach(r => {
                        const priceStr = r['Цена'] || '';
                        if (!priceStr.includes('₽')) return;
                        const priceVal = parseInt(priceStr.replace(/\D/g, ''));
                        if (!priceVal || priceVal < 5000) return;
                        if (isOneWay) {
                            allOptions.push({
                                'Цена': priceStr,
                                'Вылет': dk,
                                'Дата': r['Р1 Дата'] || '',
                                'Время': r['Р1 Вылет'] || '',
                                'Прилёт': r['Р1 Прилёт'] || '',
                                '_sort': priceVal
                            });
                        } else {
                            allOptions.push({
                                'Цена': priceStr,
                                'Вылет': dk,
                                'Ноч.': n2,
                                'Туда дата': r['Р1 Дата'] || '',
                                'Туда время': r['Р1 Вылет'] || '',
                                'Прилёт MSQ': r['Р1 Прилёт'] || '',
                                'Обратно дата': r['Р2 Дата'] || '',
                                'Обратно время': r['Р2 Вылет'] || '',
                                'Прилёт MOW': r['Р2 Прилёт'] || '',
                                '_sort': priceVal
                            });
                        }
                    });
                });
                allOptions.sort((a, b) => a._sort - b._sort);
                allOptions.forEach(r => delete r._sort);

                const optWs = XLSX.utils.json_to_sheet(allOptions);
                if (isOneWay) {
                    optWs['!cols'] = [{wch:14},{wch:8},{wch:10},{wch:8},{wch:8}];
                } else {
                    optWs['!cols'] = [{wch:14},{wch:8},{wch:5},{wch:10},{wch:8},{wch:10},{wch:12},{wch:10},{wch:10}];
                }
                XLSX.utils.book_append_sheet(summaryWb, optWs, 'Все варианты');
                console.log(`✓ "Все варианты": ${allOptions.length} рейсов по цене`);
            } catch (gridErr) {
                console.log('  (По датам: пропущено — ' + gridErr.message + ')');
            }

            const now = new Date();
            const dateStr = (now.getMonth() + 1).toString().padStart(2, '0') + '-' + now.getDate().toString().padStart(2, '0');
            const timeStr = now.toTimeString().slice(0, 5).replace(':', '-');
            const summaryFile = `output/as_СВОДКИ_${dateStr}_${timeStr}.xlsx`;
            XLSX.writeFile(summaryWb, summaryFile);
            console.log(`✓ Общая сводка: ${summaryFile}`);

            execSync(`open "${summaryFile}"`);
            console.log(`✓ Файл открыт!`);
        } else if (createdFiles.length === 1) {
            // Один файл — просто открываем
            execSync(`open "${createdFiles[0].excel}"`);
            console.log(`\n✓ Открыт файл: ${createdFiles[0].excel}`);
        }

    } catch (error) {
        console.error('\nОшибка:', error);
        try {
            await page.screenshot({ path: 'error-screenshot.png' });
            console.log('Скриншот ошибки: error-screenshot.png');
        } catch (e) {}
    } finally {
        await browser.close();
        console.log('\n✓ Браузер закрыт.\n');
    }
}

parseAviasales();
