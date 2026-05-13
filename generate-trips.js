/**
 * Генератор комбинаций дат для поиска билетов
 *
 * Использование:
 *   node generate-trips.js          — показать комбинации + команду запуска
 *   node generate-trips.js --run    — сразу запустить парсер
 *
 * Параметры — в trip-config.json
 */

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'trip-config.json'), 'utf-8'));

const route    = config['маршрут']['города'];
const names    = config['маршрут']['названия'] || route;
const pax      = config['маршрут']['пассажиры'] || 1;
const windows  = config['даты']['окна_вылета'] || [];
const nightMin = config['ночи']['мин'];
const nightMax = config['ночи']['макс'];
const stops    = config['ночи']['остановки'] || null; // для маршрутов A→B→C→A

// ─── Утилиты дат ────────────────────────────────────────────────────────────

const DAYS_IN_MONTH = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function parseDate(str) {
    const [day, month] = str.split('.').map(Number);
    return { day, month };
}

function toUrlFmt(day, month) {
    return String(day).padStart(2, '0') + String(month).padStart(2, '0');
}

function dayOfYear(day, month) {
    let n = day;
    for (let i = 1; i < month; i++) n += DAYS_IN_MONTH[i];
    return n;
}

function addDays(day, month, days) {
    let d = day + days, m = month;
    while (d > DAYS_IN_MONTH[m]) { d -= DAYS_IN_MONTH[m]; m++; if (m > 12) m = 1; }
    return { day: d, month: m };
}

function fmtDate(d) {
    return `${String(d.day).padStart(2, '0')}.${String(d.month).padStart(2, '0')}`;
}

// ─── Сборка URL ──────────────────────────────────────────────────────────────

function buildUrl(cities, dates, passengers) {
    // dates.length === cities.length - 1
    let s = '';
    for (let i = 0; i < cities.length - 1; i++) {
        s += cities[i] + toUrlFmt(dates[i].day, dates[i].month);
    }
    s += cities[cities.length - 1] + passengers;
    return `https://www.aviasales.ru/search/${s}`;
}

// ─── Генерация комбинаций ────────────────────────────────────────────────────

function generateCombinations() {
    const combinations = [];

    for (const win of windows) {
        const start   = parseDate(win['от']);
        const end     = parseDate(win['до']);
        const startN  = dayOfYear(start.day, start.month);
        const endN    = dayOfYear(end.day, end.month);

        for (let offset = 0; offset <= endN - startN; offset++) {
            const dep = addDays(start.day, start.month, offset);

            if (route.length === 2) {
                // Односторонний перелёт: A → B
                combinations.push({
                    label: `${fmtDate(dep)}  (в одну сторону)`,
                    dates: [dep],
                    url: buildUrl(route, [dep], pax)
                });

            } else if (route.length === 3) {
                // Простой туда-обратно: A → B → A
                for (let nights = nightMin; nights <= nightMax; nights++) {
                    const ret = addDays(dep.day, dep.month, nights);
                    combinations.push({
                        label: `${fmtDate(dep)} → ${fmtDate(ret)}  (${nights}н)`,
                        dates: [dep, ret],
                        url: buildUrl(route, [dep, ret], pax)
                    });
                }

            } else if (route.length === 4 && stops && stops.length >= 2) {
                // Маршрут с промежуточной остановкой: A → B → C → A
                for (let n1 = stops[0]['мин']; n1 <= stops[0]['макс']; n1++) {
                    const mid = addDays(dep.day, dep.month, n1);
                    for (let n2 = stops[1]['мин']; n2 <= stops[1]['макс']; n2++) {
                        const ret = addDays(mid.day, mid.month, n2);
                        combinations.push({
                            label: `${fmtDate(dep)} → ${fmtDate(mid)} → ${fmtDate(ret)}  (${n1}н + ${n2}н)`,
                            dates: [dep, mid, ret],
                            url: buildUrl(route, [dep, mid, ret], pax)
                        });
                    }
                }
            }
        }
    }

    return combinations;
}

// ─── Main ────────────────────────────────────────────────────────────────────

function main() {
    const routeStr = names.join(' → ');
    console.log('\n' + '█'.repeat(70));
    console.log('  ГЕНЕРАТОР КОМБИНАЦИЙ ДАТ');
    console.log('█'.repeat(70));
    console.log(`\nМаршрут:    ${routeStr}`);
    console.log(`Пассажиры:  ${pax}`);
    console.log('Окна вылета:');
    windows.forEach((w, i) => console.log(`  ${i + 1}. ${w['от']} – ${w['до']}`));
    if (route.length === 2) {
        console.log(`Тип:        в одну сторону`);
    } else if (route.length === 3) {
        console.log(`Ночей:      ${nightMin}–${nightMax}`);
    }

    const combinations = generateCombinations();

    if (combinations.length === 0) {
        console.log('\n⚠️  Нет подходящих комбинаций. Проверь параметры в trip-config.json.');
        return;
    }

    console.log(`\n✓ Сгенерировано комбинаций: ${combinations.length}\n`);
    combinations.forEach((c, i) => {
        console.log(`  ${String(i + 1).padStart(2)}. ${c.label}`);
    });

    const urlArgs = combinations.map(c => `"${c.url}"`).join(' \\\n  ');
    console.log('\n' + '─'.repeat(70));
    console.log('Команда для запуска парсера:\n');
    console.log(`node aviasales-parser.js \\\n  ${urlArgs}`);
    console.log('\n' + '─'.repeat(70));
    console.log('\nИли сразу запустить: node generate-trips.js --run\n');

    if (process.argv.includes('--run')) {
        console.log('🚀 Запускаем парсер...\n');
        spawn('node', ['aviasales-parser.js', ...combinations.map(c => c.url)], { stdio: 'inherit' });
    }
}

main();
