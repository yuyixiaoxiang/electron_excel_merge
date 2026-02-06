/**
 * 涓昏繘绋嬪叆鍙ｏ細璐熻矗鍒涘缓 Electron 绐楀彛銆佽В鏋?git/Fork 浼犲叆鐨勪笁鏂瑰悎骞跺弬鏁帮紝
 * 骞堕€氳繃 IPC 鍚戞覆鏌撹繘绋嬫彁渚?Excel 璇诲啓涓庝笁鏂?diff / merge 鐨勮兘鍔涖€?
 */
import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { spawn } from 'child_process';
import { Workbook, Worksheet, Row, Cell, CellValue } from 'exceljs';

// 淇濇寔瀵逛富绐楀彛鐨勫紩鐢紝閬垮厤琚?GC 鍥炴敹瀵艰嚧绐楀彛琚剰澶栧叧闂?
let mainWindow: BrowserWindow | null = null;

const isDev = process.env.NODE_ENV === 'development';
const DEFAULT_FROZEN_HEADER_ROWS = 3;
const DEFAULT_ROW_SIMILARITY_THRESHOLD = 0.9;
const IGNORE_BASE_IN_DIFF = true;

/**
 * CLI three-way merge arguments for git/Fork integration.
 *
 * 绾﹀畾锛堜互 Fork / git mergetool 涓轰緥锛夛細
 *   - diff 妯″紡:   app.exe OURS THEIRS
 *   - merge 妯″紡:  app.exe BASE OURS THEIRS [MERGED]
 *
 * 褰撳甫鏈?mergedPath 鏃讹紝淇濆瓨缁撴灉浼氱洿鎺ュ啓鍥?MERGED 鏂囦欢锛?
 * 鍚﹀垯浼氬洖閫€鍒拌鐩?ours锛堝綋鍓嶅垎鏀伐浣滃尯鏂囦欢锛夈€?
 */
interface CliThreeWayArgs {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mode: 'diff' | 'merge';
}

/**
 * 浠?process.argv 涓В鏋愪笁鏂瑰悎骞剁浉鍏冲弬鏁般€?
 *
 * - 寮€鍙戠幆澧冧笅 argv 褰㈠: [electron, main.js, '.', ...args]
 * - 鎵撳寘鍚?exe 涓?argv 褰㈠: [app.exe, ...args]
 */
const parseCliThreeWayArgs = (): CliThreeWayArgs | null => {
  // 瀵逛簬寮€鍙戠幆澧? process.argv = [electron, main.js, '.', ...args]
  // 瀵逛簬鎵撳寘鍚庣殑 exe: process.argv = [app.exe, ...args]
  const argStartIndex = app?.isPackaged ? 1 : 2;
  const rawArgs = process.argv.slice(argStartIndex);
  const stripOuterQuotes = (s: string) => s.replace(/^"(.*)"$/, '$1').replace(/^'(.*)'$/, '$1');
  const normalizeCliPath = (p: string) => {
    const raw = stripOuterQuotes(p);
    if (!raw) return raw;
    return path.isAbsolute(raw) ? raw : path.resolve(process.cwd(), raw);
  };
  const userArgs = rawArgs
    .map((arg) => stripOuterQuotes(arg))
    .filter((arg) => !!arg && !arg.startsWith('--'));
  // 鍏煎寮€鍙戞ā寮忎笅 `electron .` 甯︽潵鐨?app path 鍙傛暟
  if (userArgs.length >= 3) {
    const first = userArgs[0];
    const appPath = app.getAppPath ? app.getAppPath() : '';
    const firstResolved = path.resolve(first);
    const appResolved = appPath ? path.resolve(appPath) : '';
    let isDir = false;
    try {
      isDir = fs.statSync(firstResolved).isDirectory();
    } catch {
      isDir = false;
    }
    if (first === '.' || (!!appResolved && firstResolved === appResolved) || isDir) {
      userArgs.shift();
    }
  }

  // 2 涓弬鏁? 璁や负鏄?diff 妯″紡 -> base 涓?ours 鐩稿悓锛堜粎鐢ㄤ簬璁＄畻宸紓锛?
  if (userArgs.length === 2) {
    const [oursPath, theirsPath] = userArgs.map(normalizeCliPath);
    return { basePath: oursPath, oursPath, theirsPath, mode: 'diff' };
  }

  if (userArgs.length < 3) {
    return null;
  }

  const [basePath, oursPath, theirsPath, mergedPath] = userArgs.map(normalizeCliPath);
  return { basePath, oursPath, theirsPath, mergedPath, mode: 'merge' };
};

// 瑙ｆ瀽鍚姩鍙傛暟寰楀埌鐨勪笁鏂瑰悎骞朵俊鎭紙鑻ユ棤鍙傛暟鍒欎负 null锛岃蛋浜や簰寮忔ā寮忥級
const cliThreeWayArgs: CliThreeWayArgs | null = parseCliThreeWayArgs();
const getBundledGitInfo = (): { gitPath: string; env: NodeJS.ProcessEnv } | null => {
  const basePath = app?.isPackaged
    ? path.join(process.resourcesPath, 'git')
    : path.join(app.getAppPath(), 'resources', 'portable-git');
  const gitPath = path.join(basePath, 'cmd', 'git.exe');
  if (!fs.existsSync(gitPath)) return null;

  const env = { ...process.env };
  const extraPaths = [
    path.join(basePath, 'cmd'),
    path.join(basePath, 'mingw64', 'bin'),
    path.join(basePath, 'usr', 'bin'),
  ];
  const currentPath = env.PATH || env.Path || '';
  const newPath = [...extraPaths, currentPath].filter(Boolean).join(path.delimiter);
  env.PATH = newPath;
  env.Path = newPath;
  return { gitPath, env };
};

/**
 * 灏濊瘯鍦ㄧ洰鏍囨枃浠舵墍鍦ㄧ洰褰曟墽琛屼竴娆?`git add <filePath>`锛?
 * 鏂逛究鍦ㄤ綔涓?merge tool 杩愯鏃惰嚜鍔ㄦ爣璁板啿绐佸凡瑙ｅ喅銆?
 *
 * 娉ㄦ剰锛氳繖閲屽仛鐨勬槸鈥滃敖鍔涜€屼负鈥濈殑鎿嶄綔锛屽け璐ュ彧浼氭墦鍗版棩蹇楋紝涓嶄細涓柇涓绘祦绋嬨€?
 */
const gitAddFile = (filePath: string): Promise<void> => {
  return new Promise((resolve) => {
    const cwd = path.dirname(filePath);
    const gitInfo = getBundledGitInfo();
    const gitCommand = gitInfo?.gitPath ?? 'git';
    const child = spawn(gitCommand, ['add', filePath], { cwd, stdio: 'ignore', env: gitInfo?.env });

    child.on('error', (err) => {
      console.error('git add failed', err);
      resolve();
    });

    child.on('close', (code) => {
      if (code !== 0) {
        console.error('git add exited with code', code);
      }
      resolve();
    });
  });
};

/**
 * 鍒涘缓涓绘祻瑙堝櫒绐楀彛骞跺姞杞藉墠绔〉闈€?
 *
 * 寮€鍙戞ā寮忎笅杩炴帴鏈湴 webpack dev server锛?
 * 鐢熶骇妯″紡涓嬪姞杞芥墦鍖呭埌 dist 涓殑 index.html銆?
 */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:3000');
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
type SimpleCellValue = string | number | null;

interface RowRecord {
  rowNumber: number; // 1-based Excel row number
  index: number; // 0-based index in extracted rows list
  values: SimpleCellValue[];
  nonEmptyCols: number[]; // 1-based column indices with non-empty values
  key?: string | null;
}
interface ColumnTypeSignature {
  num: number;
  str: number;
  empty: number;
  other: number;
}

interface ColumnRecord {
  colNumber: number; // 1-based Excel column number
  headerText: string; // normalized header text (joined by "|")
  headerKey: string; // stronger normalized key for matching
  typeSig: ColumnTypeSignature;
  sampleValues: string[]; // normalized sample values
}

interface AlignedColumn {
  baseCol?: number | null;
  oursCol?: number | null;
  theirsCol?: number | null;
}

interface AlignedRow {
  base?: RowRecord | null;
  ours?: RowRecord | null;
  theirs?: RowRecord | null;
  key?: string | null;
  ambiguousOurs?: boolean;
  ambiguousTheirs?: boolean;
}

/**
 * 灏?ExcelJS 鐨勫鏉傚崟鍏冩牸鍊艰浆鎹负绠€鍗曞€硷紙string | number | null锛夈€?
 * 
 * ExcelJS 鐨勫崟鍏冩牸鍊煎彲鑳芥槸锛?
 * - 绠€鍗曠被鍨嬶細string銆乶umber
 * - 瀵屾枃鏈細{ richText: [{text: '...'}] }
 * - 鍏紡锛歿 formula: '...', result: value }
 * - 瓒呴摼鎺ョ瓑鍏朵粬瀵硅薄绫诲瀷
 * 
 * 璇ュ嚱鏁扮粺涓€鎻愬彇鍏朵腑鐨勫疄闄呮枃鏈?鏁板€煎唴瀹癸紝蹇界暐鏍煎紡淇℃伅銆?
 */
const getSimpleValueForMerge = (v: any): SimpleCellValue => {
  if (v === null || v === undefined) return null;
  // 澶勭悊鏃ユ湡瀵硅薄锛氳浆涓?ISO 瀛楃涓诧紝淇濇寔涓?excel:open 涓?getSimpleValue 涓€鑷?
  if (v instanceof Date) return v.toISOString();
  // 澶勭悊瀵屾枃鏈細鎷兼帴鎵€鏈夋枃鏈墖娈?
  if (typeof v === 'object' && Array.isArray((v as any).richText)) {
    const parts = (v as any).richText
      .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
      .join('');
    return parts;
  }
  // 澶勭悊瓒呴摼鎺ョ瓑鍖呭惈 text 灞炴€х殑瀵硅薄
  if (typeof v === 'object' && 'text' in v) return (v as any).text ?? null;
  // 澶勭悊鍏紡鍗曞厓鏍硷細鍙栬绠楃粨鏋?
  if (typeof v === 'object' && 'result' in v) return (v as any).result ?? null;
  // 绠€鍗曠被鍨嬬洿鎺ヨ繑鍥?
  if (typeof v === 'string' || typeof v === 'number') return v;
  // 鍏朵粬绫诲瀷杞瓧绗︿覆
  return String(v);
};

/**
 * 灏嗗崟鍏冩牸鍊兼爣鍑嗗寲涓哄瓧绗︿覆锛岀敤浜庢瘮杈冨拰鏄剧ず銆?
 * - null/undefined 鈫?绌哄瓧绗︿覆
 * - 瀛楃涓?鈫?鍘婚櫎棣栧熬绌烘牸
 * - 鏁板瓧 鈫?杞瓧绗︿覆
 */
const normalizeCellValue = (v: SimpleCellValue): string => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  return String(v);
};

/**
 * 鏍囧噯鍖栦富閿垪鐨勫€硷紝鐢ㄤ簬琛屽榻愩€?
 * 绌哄瓧绗︿覆瑙嗕负 null锛堝嵆鏃犱富閿級锛屾柟渚垮悗缁垽鏂€?
 */
const normalizeKeyValue = (v: SimpleCellValue): string | null => {
  const s = normalizeCellValue(v);
  return s === '' ? null : s;
};

/**
 * 鏍囧噯鍖栬〃澶存枃鏈紝鐢ㄤ簬鍒楀尮閰嶃€?
 * 杞负灏忓啓浠ュ拷鐣ュぇ灏忓啓宸紓銆?
 */
const normalizeHeaderText = (v: SimpleCellValue): string => {
  const s = normalizeCellValue(v);
  if (!s) return '';
  return s.toLowerCase();
};
/**
 * 鐢熸垚鏇村己鐨勮〃澶村尮閰嶉敭锛岀敤浜庣簿纭尮閰嶅垪銆?
 * - 杞皬鍐?
 * - 鍘婚櫎鎵€鏈夌┖鐧?
 * - 鍙繚鐣欏瓧姣嶃€佹暟瀛椼€佷腑鏂囧瓧绗?
 * 
 * 渚嬪锛?Icon鍚嶇О, Asset..." 鈫?"icon鍚嶇Оasset"
 * 杩欐牱鍗充娇鏍煎紡鐣ユ湁涓嶅悓锛屼篃鑳藉尮閰嶄笂鐩稿悓璇箟鐨勫垪銆?
 */
const normalizeHeaderKey = (text: string): string => {
  if (!text) return '';
  return text
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^0-9a-z\u4e00-\u9fa5]/gi, '');
};

/**
 * 涓哄伐浣滆〃鐨勬瘡涓€鍒楁彁鍙栫壒寰佷俊鎭紝鐢ㄤ簬鍒楀榻愮畻娉曘€?
 * 
 * @param ws ExcelJS 宸ヤ綔琛ㄥ璞?
 * @param headerCount 琛ㄥご琛屾暟锛堝墠N琛岃涓鸿〃澶达級
 * @param sampleRows 閲囨牱琛屾暟锛堢敤浜庣被鍨嬪拰鏍锋湰鍊肩粺璁★級
 * @returns 鍒楃壒寰佽褰曟暟缁?
 * 
 * 鐗瑰緛鍖呮嫭锛?
 * 1. headerText: 琛ㄥご鏂囨湰锛堝琛岀敤 | 鍒嗛殧锛?
 * 2. headerKey: 鏍囧噯鍖栫殑琛ㄥご閿紙鐢ㄤ簬绮剧‘鍖归厤锛?
 * 3. typeSig: 鏁版嵁绫诲瀷绛惧悕锛坣um/str/empty/other 鐨勫垎甯冿級
 * 4. sampleValues: 鏍锋湰鍊奸泦鍚堬紙鐢ㄤ簬鍐呭鐩镐技搴︽瘮杈冿級
 * 
 * 娉ㄦ剰锛氬畬鍏ㄧ┖鐨勫垪锛堣〃澶村拰鏁版嵁閮戒负绌猴級浼氳璺宠繃锛屼笉鐢熸垚璁板綍銆?
 */
const buildColumnRecords = (
  ws: any,
  headerCount: number,
  sampleRows: number,
): ColumnRecord[] => {
  if (!ws) return [];
  // 鑾峰彇宸ヤ綔琛ㄥ疄闄呭垪鏁?
  const actualColCount = Math.max(ws?.actualColumnCount ?? 0, ws?.columnCount ?? 0);
  const maxRow = Math.max(ws?.actualRowCount ?? 0, ws?.rowCount ?? 0, headerCount);
  const records: ColumnRecord[] = [];
  
  // 閬嶅巻姣忎竴鍒?
  for (let col = 1; col <= actualColCount; col += 1) {
    // 1. 鎻愬彇琛ㄥご鏂囨湰锛堟嫾鎺ュ墠 headerCount 琛岋級
    const headerParts: string[] = [];
    for (let r = 1; r <= headerCount; r += 1) {
      const row = ws.getRow(r);
      const raw = getSimpleValueForMerge(row.getCell(col)?.value);
      const text = normalizeHeaderText(raw);
      if (text) headerParts.push(text);
    }
  const headerText = headerParts.join('|');
  const headerKey = normalizeHeaderKey(headerText);
    const typeSig: ColumnTypeSignature = { num: 0, str: 0, empty: 0, other: 0 };
    const sampleSet = new Set<string>();
    let sampled = 0;
    for (let r = headerCount + 1; r <= maxRow && sampled < sampleRows; r += 1) {
      const row = ws.getRow(r);
      const raw = getSimpleValueForMerge(row.getCell(col)?.value);
      const norm = normalizeCellValue(raw);
      if (norm === '') {
        typeSig.empty += 1;
        sampled += 1;
        continue;
      }
      if (typeof raw === 'number') typeSig.num += 1;
      else if (typeof raw === 'string') typeSig.str += 1;
      else typeSig.other += 1;
      sampleSet.add(norm);
      sampled += 1;
    }
    const sampleValues = Array.from(sampleSet).slice(0, 12);
    const hasDataSample = sampleValues.length > 0 || typeSig.num > 0 || typeSig.str > 0 || typeSig.other > 0;
    const isFullyEmpty = !headerText && !hasDataSample;
    if (isFullyEmpty) continue;

    records.push({
      colNumber: col,
      headerText,
      headerKey,
      typeSig,
      sampleValues,
    });
  }
  return records;
};

/**
 * 璁＄畻涓や釜瀛楃涓茬殑鐩镐技搴︼紙浣跨敤 Levenshtein 璺濈锛夈€?
 * 
 * @returns 0-1 涔嬮棿鐨勭浉浼煎害锛? 琛ㄧず瀹屽叏鐩稿悓锛? 琛ㄧず瀹屽叏涓嶅悓銆?
 * 
 * 绠楁硶锛歀evenshtein 璺嶇绠楁硶锛堝姩鎬佽鍒掞級
 * - 璁＄畻灏嗗瓧绗︿覆 a 杞崲涓?b 鎵€闇€鐨勬渶灏忕紪杈戞楠わ紙鎻掑叆銆佸垹闄ゃ€佹浛鎹級
 * - 鐩镐技搴?= 1 - (璺嶇 / 杈冮暱瀛楃涓查暱搴?
 */
const stringSimilarity = (a: string, b: string): number => {
  if (!a && !b) return 1;
  if (!a || !b) return 0;
  const s = a.toLowerCase();
  const t = b.toLowerCase();
  if (s === t) return 1;
  const n = s.length;
  const m = t.length;
  if (n === 0 || m === 0) return 0;
  // 鍔ㄦ€佽鍒掕绠楃紪杈戣窛绂?
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  // 鍒濆鍖栵細绗琲涓瓧绗﹁浆鎹负绌洪渶瑕乮姝?
  for (let i = 0; i <= n; i += 1) dp[i][0] = i;
  for (let j = 0; j <= m; j += 1) dp[0][j] = j;
  // 濉〃锛氳绠楁瘡涓瓙闂鐨勬渶灏忕紪杈戣窛绂?
  for (let i = 1; i <= n; i += 1) {
    for (let j = 1; j <= m; j += 1) {
      const cost = s[i - 1] === t[j - 1] ? 0 : 1;  // 瀛楃鐩稿悓鏃犻渶鏇挎崲
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,       // 鍒犻櫎
        dp[i][j - 1] + 1,       // 鎻掑叆
        dp[i - 1][j - 1] + cost, // 鏇挎崲
      );
    }
  }
  const dist = dp[n][m];
  // 褰掍竴鍖栦负 0-1 涔嬮棿鐨勭浉浼煎害
  return 1 - dist / Math.max(n, m);
};

/**
 * 璁＄畻涓や釜鍒楃殑鏁版嵁绫诲瀷绛惧悕鐩镐技搴︺€?
 * 
 * 绫诲瀷绛惧悕 = { num, str, empty, other } 鐨勫垎甯冩瘮渚嬨€?
 * 鐩镐技搴?= 1 - (姣斾緥宸紓鐨勬€诲拰 / 2)銆?
 * 
 * 渚嬪锛?
 * - A鍒楋細80% 鏁板瓧锛?0% 瀛楃涓?
 * - B鍒楋細85% 鏁板瓧锛?5% 瀛楃涓?
 * - 鐩镐技搴﹀緢楂橈紝寰堝彲鑳芥槸鍚屼竴鍒?
 */
const typeSignatureSimilarity = (a: ColumnTypeSignature, b: ColumnTypeSignature): number => {
  const totalA = a.num + a.str + a.empty + a.other;
  const totalB = b.num + b.str + b.empty + b.other;
  if (totalA === 0 && totalB === 0) return 1;
  if (totalA === 0 || totalB === 0) return 0;
  const pa = {
    num: a.num / totalA,
    str: a.str / totalA,
    empty: a.empty / totalA,
    other: a.other / totalA,
  };
  const pb = {
    num: b.num / totalB,
    str: b.str / totalB,
    empty: b.empty / totalB,
    other: b.other / totalB,
  };
  const dist =
    Math.abs(pa.num - pb.num) +
    Math.abs(pa.str - pb.str) +
    Math.abs(pa.empty - pb.empty) +
    Math.abs(pa.other - pb.other);
  return 1 - dist / 2;
};

const valueSimilarity = (a: string[], b: string[]): number => {
  if (a.length === 0 && b.length === 0) return 1;
  if (a.length === 0 || b.length === 0) return 0;
  const setA = new Set(a);
  const setB = new Set(b);
  let intersect = 0;
  setA.forEach((v) => {
    if (setB.has(v)) intersect += 1;
  });
  const union = setA.size + setB.size - intersect;
  if (union === 0) return 0;
  return intersect / union;
};

const columnSimilarity = (a: ColumnRecord, b: ColumnRecord): number => {
  const headerSim = stringSimilarity(a.headerKey || a.headerText, b.headerKey || b.headerText);
  const typeSim = typeSignatureSimilarity(a.typeSig, b.typeSig);
  const valSim = valueSimilarity(a.sampleValues, b.sampleValues);
  const hasHeader = (a.headerKey || a.headerText) && (b.headerKey || b.headerText);
  const wHeader = hasHeader ? 0.6 : 0.2;
  const wType = 0.2;
  const wVal = 0.2;
  const sum = wHeader + wType + wVal;
  return (wHeader * headerSim + wType * typeSim + wVal * valSim) / sum;
};

const alignColumnsBySimilarity = (
  baseCols: ColumnRecord[],
  sideCols: ColumnRecord[],
): { matched: Map<number, number>; gaps: Map<number, ColumnRecord[]> } => {
  const baseTokens = baseCols.map((c, i) => (c.headerKey || c.headerText ? (c.headerKey || c.headerText) : `__EMPTY_${i}`));
  const sideTokens = sideCols.map((c, i) => (c.headerKey || c.headerText ? (c.headerKey || c.headerText) : `__EMPTY_${i}`));
  const anchorPairs = lcsMatchPairs(baseTokens, sideTokens);
  const matched = new Map<number, number>();
  const usedSide = new Set<number>();
  for (const p of anchorPairs) {
    matched.set(p.aIndex, p.bIndex);
    usedSide.add(p.bIndex);
  }

  anchorPairs.sort((a, b) => a.aIndex - b.aIndex);

  const threshold = 0.55;
  const headerThreshold = 0.8;
  const matchSegment = (baseIdxs: number[], sideIdxs: number[]) => {
    if (baseIdxs.length === 0 || sideIdxs.length === 0) return;
    const pairs: Array<{ b: number; s: number; score: number }> = [];
    for (const b of baseIdxs) {
      for (const s of sideIdxs) {
        const headerA = baseCols[b].headerKey || baseCols[b].headerText;
        const headerB = sideCols[s].headerKey || sideCols[s].headerText;
        const headerSim = stringSimilarity(headerA, headerB);
        if (headerA && headerB && headerSim < headerThreshold) continue;
        const score = columnSimilarity(baseCols[b], sideCols[s]);
        if (score >= threshold) pairs.push({ b, s, score });
      }
    }
    pairs.sort((a, b) => b.score - a.score);
    for (const p of pairs) {
      if (matched.has(p.b)) continue;
      if (usedSide.has(p.s)) continue;
      matched.set(p.b, p.s);
      usedSide.add(p.s);
    }
  };

  let prevBase = -1;
  let prevSide = -1;
  for (const anchor of anchorPairs) {
    const baseIdxs: number[] = [];
    const sideIdxs: number[] = [];
    for (let b = prevBase + 1; b < anchor.aIndex; b += 1) baseIdxs.push(b);
    for (let s = prevSide + 1; s < anchor.bIndex; s += 1) sideIdxs.push(s);
    matchSegment(baseIdxs, sideIdxs);
    prevBase = anchor.aIndex;
    prevSide = anchor.bIndex;
  }
  if (prevBase < baseCols.length - 1 || prevSide < sideCols.length - 1) {
    const baseIdxs: number[] = [];
    const sideIdxs: number[] = [];
    for (let b = prevBase + 1; b < baseCols.length; b += 1) baseIdxs.push(b);
    for (let s = prevSide + 1; s < sideCols.length; s += 1) sideIdxs.push(s);
    matchSegment(baseIdxs, sideIdxs);
  }

  const gaps = new Map<number, ColumnRecord[]>();
  const matchedPairsBySide = Array.from(matched.entries())
    .map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }))
    .sort((a, b) => a.sideIndex - b.sideIndex);
  for (let s = 0; s < sideCols.length; s += 1) {
    if (usedSide.has(s)) continue;
    let gap = -1;
    for (const p of matchedPairsBySide) {
      if (p.sideIndex < s) gap = p.baseIndex;
      if (p.sideIndex >= s) break;
    }
    if (!gaps.has(gap)) gaps.set(gap, []);
    gaps.get(gap)!.push(sideCols[s]);
  }

  return { matched, gaps };
};

const buildAlignedColumns = (
  baseWs: any,
  oursWs: any,
  theirsWs: any,
  headerCount: number,
): AlignedColumn[] => {
  const sampleRows = 20;
  const baseCols = buildColumnRecords(baseWs, headerCount, sampleRows);
  const oursCols = buildColumnRecords(oursWs, headerCount, sampleRows);
  const theirsCols = buildColumnRecords(theirsWs, headerCount, sampleRows);

  const alignBase = baseCols.length > 0 ? baseCols : oursCols.length > 0 ? oursCols : theirsCols;
  const baseRefCols = alignBase;
  const oursAlign = alignColumnsBySimilarity(baseRefCols, oursCols);
  const theirsAlign = alignColumnsBySimilarity(baseRefCols, theirsCols);

  const aligned: AlignedColumn[] = [];
  const addGapCols = (gapIndex: number) => {
    const oursGap = oursAlign.gaps.get(gapIndex) ?? [];
    const theirsGap = theirsAlign.gaps.get(gapIndex) ?? [];
    for (const c of oursGap) aligned.push({ oursCol: c.colNumber ?? null });
    for (const c of theirsGap) aligned.push({ theirsCol: c.colNumber ?? null });
  };

  addGapCols(-1);
  for (let i = 0; i < baseRefCols.length; i += 1) {
    const baseColNumber = baseRefCols[i]?.colNumber ?? null;
    const oursIndex = oursAlign.matched.get(i);
    const theirsIndex = theirsAlign.matched.get(i);
    aligned.push({
      baseCol: baseColNumber,
      oursCol: typeof oursIndex === 'number' ? oursCols[oursIndex]?.colNumber ?? null : null,
      theirsCol: typeof theirsIndex === 'number' ? theirsCols[theirsIndex]?.colNumber ?? null : null,
    });
    addGapCols(i);
  }

  if (baseRefCols.length === 0) {
    // base/ours 氇憪涓虹┖鏃讹紝鐩存帴鎸?theirs 杩藉姞
    for (const c of theirsCols) {
      aligned.push({ theirsCol: c.colNumber ?? null });
    }
  }

  return aligned;
};

const colNumberToLabel = (colNumber: number): string => {
  let n = Math.max(1, Math.floor(colNumber));
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode('A'.charCodeAt(0) + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
};

/**
 * 浠庡伐浣滆〃涓彁鍙栬璁板綍锛堟湭瀵归綈鐗堟湰锛岀敤浜庡崟鏂囦欢鎴栧垪瀵归綈鍓嶏級銆?
 * 
 * @param ws ExcelJS 宸ヤ綔琛ㄥ璞?
 * @param colCount 鍒楁暟
 * @param primaryKeyCol 涓婚敭鍒楀彿锛?-based锛?1 琛ㄧず鏃犱富閿級
 * @returns 琛岃褰曟暟缁勶紝姣忔潯璁板綍鍖呭惈锛?
 *   - rowNumber: Excel 涓殑鍘熷琛屽彿
 *   - index: 鍦ㄦ彁鍙栧垪琛ㄤ腑鐨勭储寮?
 *   - values: 鎵€鏈夊垪鐨勫€兼暟缁?
 *   - nonEmptyCols: 闈炵┖鍒楃殑鍒楀彿鍒楄〃
 *   - key: 涓婚敭鍊硷紙濡傛灉鏈夛級
 * 
 * 娉ㄦ剰锛氬畬鍏ㄧ┖鐨勮浼氳璺宠繃銆?
 */
const buildRowRecords = (ws: any, colCount: number, primaryKeyCol: number): RowRecord[] => {
  const rows: RowRecord[] = [];
  let index = 0;
  // 閬嶅巻鎵€鏈夐潪绌鸿
  ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    const values: SimpleCellValue[] = [];
    const nonEmptyCols: number[] = [];
    // 璇诲彇姣忎竴鍒楃殑鍊?
    for (let col = 1; col <= colCount; col += 1) {
      const cell = row.getCell(col);
      const value = getSimpleValueForMerge(cell?.value);
      values.push(value);
      if (value !== null && value !== '') {
        nonEmptyCols.push(col);
      }
    }
    // 璺宠繃瀹屽叏绌虹殑琛?
    if (nonEmptyCols.length === 0) return;
    // 鎻愬彇涓婚敭鍊硷紙濡傛灉鏈夋寚瀹氫富閿垪锛?
    const key =
      primaryKeyCol >= 1 && primaryKeyCol <= colCount
        ? normalizeKeyValue(values[primaryKeyCol - 1])
        : null;
    rows.push({ rowNumber, index, values, nonEmptyCols, key });
    index += 1;
  });
  return rows;
};

const buildHeaderRowRecord = (ws: any, rowNumber: number, colCount: number, primaryKeyCol: number): RowRecord => {
  const values: SimpleCellValue[] = [];
  const nonEmptyCols: number[] = [];
  const row = ws.getRow(rowNumber);
  for (let col = 1; col <= colCount; col += 1) {
    const cell = row.getCell(col);
    const value = getSimpleValueForMerge(cell?.value);
    values.push(value);
    if (value !== null && value !== '') {
      nonEmptyCols.push(col);
    }
  }
  const key =
    primaryKeyCol >= 1 && primaryKeyCol <= colCount
      ? normalizeKeyValue(values[primaryKeyCol - 1])
      : null;
  return {
    rowNumber,
    index: rowNumber - 1,
    values,
    nonEmptyCols,
    key,
  };
};

/**
 * 浠庡伐浣滆〃涓彁鍙栬璁板綍锛堝垪瀵归綈鐗堟湰锛夈€?
 * 
 * 涓?buildRowRecords 鐨勫尯鍒細
 * - 浣跨敤瀵归綈鍚庣殑鍒楅『搴?
 * - 鏍规嵁 side 鍙傛暟浠庡搴旂殑鐗╃悊鍒楄鍙栧€?
 * - 濡傛灉鏌愪竴鍒楀湪璇?side 涓嶅瓨鍦紝瀵瑰簲浣嶇疆濉?null
 * 
 * @param alignedColumns 瀵归綈鍚庣殑鍒楀厓淇℃伅
 * @param primaryKeyColAligned 涓婚敭鍒楀湪瀵归綈鍚庡簭鍒椾腑鐨勪綅缃?
 * @param side 褰撳墠澶勭悊鐨勬槸 base/ours/theirs 鍝竴渚?
 * 
 * 渚嬪锛?
 * - alignedColumns[2] = { baseCol: 3, oursCol: null, theirsCol: 2 }
 * - 瀵逛簬 ours 渚э紝绗?涓榻愬垪鐨勫€间細鏄?null锛堝洜涓?ours 娌℃湁杩欎竴鍒楋級
 */
const buildRowRecordsAligned = (
  ws: any,
  alignedColumns: AlignedColumn[],
  primaryKeyColAligned: number,
  side: 'base' | 'ours' | 'theirs',
): RowRecord[] => {
  const rows: RowRecord[] = [];
  let index = 0;
  ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    const values: SimpleCellValue[] = [];
    const nonEmptyCols: number[] = [];
    // 鎸夌収瀵归綈鍚庣殑鍒楅『搴忚鍙栧€?
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const colMeta = alignedColumns[i];
      // 鏍规嵁 side 鑾峰彇瀵瑰簲鐨勭墿鐞嗗垪鍙?
      const colNumber =
        side === 'base' ? colMeta.baseCol : side === 'ours' ? colMeta.oursCol : colMeta.theirsCol;
      let value: SimpleCellValue = null;
      // 濡傛灉璇?side 鏈夎繖涓€鍒楋紝鍒欒鍙栧€硷紱鍚﹀垯涓?null
      if (colNumber) {
        const cell = row.getCell(colNumber);
        value = getSimpleValueForMerge(cell?.value);
      }
      values.push(value);
      if (value !== null && value !== '') nonEmptyCols.push(i + 1);
    }
    if (nonEmptyCols.length === 0) return;
    const key =
      primaryKeyColAligned >= 1 && primaryKeyColAligned <= alignedColumns.length
        ? normalizeKeyValue(values[primaryKeyColAligned - 1])
        : null;
    rows.push({ rowNumber, index, values, nonEmptyCols, key });
    index += 1;
  });
  return rows;
};

const buildHeaderRowRecordAligned = (
  ws: any,
  rowNumber: number,
  alignedColumns: AlignedColumn[],
  primaryKeyColAligned: number,
  side: 'base' | 'ours' | 'theirs',
): RowRecord => {
  const values: SimpleCellValue[] = [];
  const nonEmptyCols: number[] = [];
  const row = ws.getRow(rowNumber);
  for (let i = 0; i < alignedColumns.length; i += 1) {
    const colMeta = alignedColumns[i];
    const colNumber =
      side === 'base' ? colMeta.baseCol : side === 'ours' ? colMeta.oursCol : colMeta.theirsCol;
    let value: SimpleCellValue = null;
    if (colNumber) {
      const cell = row.getCell(colNumber);
      value = getSimpleValueForMerge(cell?.value);
    }
    values.push(value);
    if (value !== null && value !== '') nonEmptyCols.push(i + 1);
  }
  const key =
    primaryKeyColAligned >= 1 && primaryKeyColAligned <= alignedColumns.length
      ? normalizeKeyValue(values[primaryKeyColAligned - 1])
      : null;
  return {
    rowNumber,
    index: rowNumber - 1,
    values,
    nonEmptyCols,
    key,
  };
};

/**
 * 鍒ゆ柇涓よ鏄惁瀹屽叏鐩哥瓑銆?
 * 
 * 鐩哥瓑鐨勫畾涔夛細鎵€鏈夐潪绌哄垪鐨勫€煎畬鍏ㄧ浉鍚屻€?
 * 鍙瘮杈冧袱琛屼腑鑷冲皯鏈変竴琛岄潪绌虹殑鍒椼€?
 */
const rowsEqual = (a: RowRecord, b: RowRecord): boolean => {
  // 鏀堕泦涓よ鐨勬墍鏈夐潪绌哄垪
  const cols = new Set<number>();
  a.nonEmptyCols.forEach((c) => cols.add(c));
  b.nonEmptyCols.forEach((c) => cols.add(c));
  // 閫愬垪姣旇緝
  for (const col of cols) {
    const av = normalizeCellValue(a.values[col - 1] ?? null);
    const bv = normalizeCellValue(b.values[col - 1] ?? null);
    if (av !== bv) return false;
  }
  return true;
};

/**
 * 璁＄畻涓よ鐨勭浉浼煎害銆?
 * 
 * @returns 0-1 涔嬮棿鐨勭浉浼煎害锛? 琛ㄧず瀹屽叏鐩稿悓銆?
 * 
 * 绠楁硶锛?
 * 1. 鏀堕泦涓よ鐨勬墍鏈夐潪绌哄垪
 * 2. 璁＄畻鐩稿悓鍊肩殑鍒楁暟 / 鎬诲垪鏁?
 * 3. 璺宠繃涓よ竟閮戒负绌虹殑鍒楋紙涓嶈鍏ユ€绘暟锛?
 * 
 * 渚嬪锛?
 * - A琛? [1, "abc", null, "xyz"]
 * - B琛? [1, "abc", "new", "xyz"]
 * - 鐩镐技搴?= 3/4 = 0.75锛堢3鍒椾笉鍚岋級
 */
const rowSimilarity = (a: RowRecord, b: RowRecord): number => {
  const cols = new Set<number>();
  a.nonEmptyCols.forEach((c) => cols.add(c));
  b.nonEmptyCols.forEach((c) => cols.add(c));
  if (cols.size === 0) return 1;
  let same = 0;
  let total = 0;
  for (const col of cols) {
    const av = normalizeCellValue(a.values[col - 1] ?? null);
    const bv = normalizeCellValue(b.values[col - 1] ?? null);
    // 璺宠繃涓よ竟閮戒负绌虹殑鍒?
    if (av === '' && bv === '') continue;
    total += 1;
    if (av === bv) same += 1;
  }
  if (total === 0) return 1;
  return same / total;
};

/**
 * 璁＄畻琛岀殑鐘舵€侊紙鍩轰簬涓夋柟瀵规瘮锛夈€?
 * 
 * @returns 琛岀姸鎬侊細
 *   - 'ambiguous': 鍖归厤鏈夋涔夛紙澶氫釜鍊欓€夎锛?
 *   - 'added': 鏂板琛岋紙base 娌℃湁锛宻ide 鏈夛級
 *   - 'deleted': 鍒犻櫎琛岋紙base 鏈夛紝side 娌℃湁锛?
 *   - 'unchanged': 鏈彉鍖栵紙鍐呭瀹屽叏鐩稿悓锛?
 *   - 'modified': 淇敼琛岋紙鍐呭涓嶅悓锛?
 */
const computeRowStatus = (
  baseRow: RowRecord | null | undefined,
  sideRow: RowRecord | null | undefined,
  isAmbiguous: boolean | undefined,
): RowStatus => {
  if (isAmbiguous) return 'ambiguous';
  if (!baseRow && sideRow) return 'added';
  if (baseRow && !sideRow) return 'deleted';
  if (!baseRow && !sideRow) return 'unchanged';
  if (baseRow && sideRow && rowsEqual(baseRow, sideRow)) return 'unchanged';
  return 'modified';
};

const makeAddress = (col: number, row: number): string => {
  return `${colNumberToLabel(col)}${row}`;
};

const estimateSideIndex = (
  baseIndex: number,
  matchedPairs: Array<{ baseIndex: number; sideIndex: number }>,
): number => {
  if (matchedPairs.length === 0) return baseIndex;
  let prev: { baseIndex: number; sideIndex: number } | null = null;
  let next: { baseIndex: number; sideIndex: number } | null = null;
  for (const p of matchedPairs) {
    if (p.baseIndex < baseIndex) prev = p;
    if (p.baseIndex > baseIndex) {
      next = p;
      break;
    }
  }
  if (prev && next) {
    const t = (baseIndex - prev.baseIndex) / Math.max(1, next.baseIndex - prev.baseIndex);
    return Math.round(prev.sideIndex + t * (next.sideIndex - prev.sideIndex));
  }
  if (prev) return prev.sideIndex + (baseIndex - prev.baseIndex);
  if (next) return next.sideIndex - (next.baseIndex - baseIndex);
  return baseIndex;
};

type DiffOp =
  | { type: 'equal'; aIndex: number; bIndex: number }
  | { type: 'delete'; aIndex: number }
  | { type: 'insert'; bIndex: number };
/**
 * 璁＄畻鏈€闀垮叕鍏卞瓙搴忓垪锛圠CS锛夊苟杩斿洖鍖归厤瀵广€?
 * 
 * 鐢ㄤ簬鍒?琛屽榻愮殑閿佺偣鍖归厤锛氭壘鍒颁袱涓簭鍒椾腑纭畾鐩稿悓鐨勫厓绱犱綔涓衡€滈攣鐐光€濄€?
 * 
 * @param a 绗竴涓瓧绗︿覆鏁扮粍
 * @param b 绗簩涓瓧绗︿覆鏁扮粍
 * @returns 鍖归厤瀵规暟缁勶紝鎸夌収鍑虹幇椤哄簭鎺掑垪
 * 
 * 渚嬪锛?
 * - a = ["A", "B", "C", "D"]
 * - b = ["A", "X", "B", "D"]
 * - 杩斿洖: [{ aIndex: 0, bIndex: 0 }, { aIndex: 1, bIndex: 2 }, { aIndex: 3, bIndex: 3 }]
 * - 鍗?A, B, D 涓変釜鍏冪礌鏄叕鍏辩殑
 * 
 * 绠楁硶锛氬姩鎬佽鍒?+ 鍥炴函
 * - dp[i][j] = a[0..i-1] 鍜?b[0..j-1] 鐨?LCS 闀垮害
 * - 鍥炴函鎵惧埌瀹為檯鍖归厤鐨勪綅缃?
 */
const lcsMatchPairs = (a: string[], b: string[]): Array<{ aIndex: number; bIndex: number }> => {
  const n = a.length;
  const m = b.length;
  // 鍔ㄦ€佽鍒掕〃锛歞p[i][j] = LCS 闀垮害
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  // 濉〃锛氳绠?LCS 闀垮害
  for (let i = 1; i <= n; i += 1) {
    for (let j = 1; j <= m; j += 1) {
      if (a[i - 1] === b[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;  // 鍖归厤锛岄暱搴?1
      else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);        // 涓嶅尮閰嶏紝鍙栨渶澶у€?
    }
  }
  // 鍥炴函锛氫粠 dp 琛ㄤ腑鎻愬彇瀹為檯鍖归厤瀵?
  const pairs: Array<{ aIndex: number; bIndex: number }> = [];
  let i = n;
  let j = m;
  while (i > 0 && j > 0) {
    if (a[i - 1] === b[j - 1]) {
      // 褰撳墠鍏冪礌鍖归厤锛岃褰曞苟缁х画鍥炴函
      pairs.push({ aIndex: i - 1, bIndex: j - 1 });
      i -= 1;
      j -= 1;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      i -= 1;  // 鍚戜笂鍥炴函
    } else {
      j -= 1;  // 鍚戝乏鍥炴函
    }
  }
  // 鍥炴函鏄粠鍚庡線鍓嶏紝闇€瑕佸弽杞?
  return pairs.reverse();
};

const myersDiff = (a: string[], b: string[]): DiffOp[] => {
  const n = a.length;
  const m = b.length;
  const max = n + m;
  let v = new Map<number, number>();
  v.set(1, 0);
  const trace: Map<number, number>[] = [];

  for (let d = 0; d <= max; d += 1) {
    const vSnap = new Map<number, number>();
    for (let k = -d; k <= d; k += 2) {
      let x: number;
      if (k === -d || (k !== d && (v.get(k - 1) ?? 0) < (v.get(k + 1) ?? 0))) {
        x = v.get(k + 1) ?? 0;
      } else {
        x = (v.get(k - 1) ?? 0) + 1;
      }
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) {
        x += 1;
        y += 1;
      }
      vSnap.set(k, x);
      if (x >= n && y >= m) {
        trace.push(vSnap);
        // backtrack
        const ops: DiffOp[] = [];
        let x2 = n;
        let y2 = m;
        for (let d2 = trace.length - 1; d2 >= 0; d2 -= 1) {
          const v2 = trace[d2];
          const k2 = x2 - y2;
          let prevK: number;
          if (k2 === -d2 || (k2 !== d2 && (v2.get(k2 - 1) ?? 0) < (v2.get(k2 + 1) ?? 0))) {
            prevK = k2 + 1;
          } else {
            prevK = k2 - 1;
          }
          const prevX = v2.get(prevK) ?? 0;
          const prevY = prevX - prevK;
          while (x2 > prevX && y2 > prevY) {
            ops.push({ type: 'equal', aIndex: x2 - 1, bIndex: y2 - 1 });
            x2 -= 1;
            y2 -= 1;
          }
          if (d2 === 0) break;
          if (x2 === prevX) {
            ops.push({ type: 'insert', bIndex: y2 - 1 });
            y2 -= 1;
          } else {
            ops.push({ type: 'delete', aIndex: x2 - 1 });
            x2 -= 1;
          }
        }
        return ops.reverse();
      }
    }
    trace.push(vSnap);
    v = vSnap;
  }
  return [];
};

/**
 * 鍩轰簬涓婚敭鍒楀榻愯銆?
 * 
 * 杩欐槸琛屽榻愮殑涓昏鏂规硶锛岄€傜敤浜庢湁鍞竴鏍囪瘑鍒楋紙濡?ID锛夌殑鏁版嵁銆?
 * 
 * @param baseRows base 鐨勮璁板綍
 * @param oursRows ours 鐨勮璁板綍
 * @param theirsRows theirs 鐨勮璁板綍
 * @param keyCol 涓婚敭鍒楀彿锛?-based锛?
 * @param rowSimilarityThreshold 鐩镐技搴﹂槇鍊硷紙鐢ㄤ簬姝т箟妫€娴嬶級
 * @returns 瀵归綈缁撴灉 + 姝т箟琛岄泦鍚?
 * 
 * 绠楁硶姝ラ锛?
 * 1. 鎸変富閿€煎垎缁勶細Map<key, RowRecord[]>
 * 2. 瀵规瘡涓富閿€硷細
 *    - 濡傛灉 base/ours/theirs 閮芥湁涓旀瘡渚у彧鏈?1 鏉?鈫?鐩存帴鍖归厤
 *    - 濡傛灉鏌愪晶鏈夊鏉＄浉鍚屼富閿?鈫?妫€娴嬫涔夛紙鐩镐技搴﹀尮閰嶏級
 * 3. 杩斿洖瀵归綈鍚庣殑涓夊厓缁勶細(base, ours, theirs)
 * 
 * 姝т箟鍦烘櫙锛?
 * - 涓婚敭鍊肩浉鍚屼絾鍏朵粬鍒楀唴瀹逛笉鍚岀殑澶氳
 * - 姝ゆ椂鏃犳硶纭畾鍝竴琛屽搴斿摢涓€琛岋紝鏍囪涓?ambiguous
 */
const alignRowsByKey = (
  baseRows: RowRecord[],
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
  keyCol: number,
  rowSimilarityThreshold: number,
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  const groupByKey = (rows: RowRecord[]) => {
    const m = new Map<string, RowRecord[]>();
    rows.forEach((r) => {
      if (!r.key) return;
      if (!m.has(r.key)) m.set(r.key, []);
      m.get(r.key)!.push(r);
    });
    return m;
  };
  const rowSimilarityIgnoringKey = (a: RowRecord, b: RowRecord): number => {
    if (keyCol < 1) return rowSimilarity(a, b);
    const cols = new Set<number>();
    a.nonEmptyCols.forEach((c) => cols.add(c));
    b.nonEmptyCols.forEach((c) => cols.add(c));
    if (cols.size === 0) return 1;
    let same = 0;
    let total = 0;
    for (const col of cols) {
      if (col === keyCol) continue;
      const av = normalizeCellValue(a.values[col - 1] ?? null);
      const bv = normalizeCellValue(b.values[col - 1] ?? null);
      if (av === '' && bv === '') continue;
      total += 1;
      if (av === bv) same += 1;
    }
    if (total === 0) return 1;
    return same / total;
  };

  const baseByKeyList = groupByKey(baseRows);
  const oursByKeyList = groupByKey(oursRows);
  const theirsByKeyList = groupByKey(theirsRows);

  const baseCounts = new Map<string, number>();
  baseByKeyList.forEach((list, key) => baseCounts.set(key, list.length));
  const oursCounts = new Map<string, number>();
  oursByKeyList.forEach((list, key) => oursCounts.set(key, list.length));
  const theirsCounts = new Map<string, number>();
  theirsByKeyList.forEach((list, key) => theirsCounts.set(key, list.length));

  const occurrenceIndex = (rows: RowRecord[]) => {
    const occ = new Map<number, number>();
    const counters = new Map<string, number>();
    rows.forEach((r) => {
      if (!r.key) return;
      const next = (counters.get(r.key) ?? 0) + 1;
      counters.set(r.key, next);
      occ.set(r.index, next - 1);
    });
    return occ;
  };

  const baseOcc = occurrenceIndex(baseRows);

  const matchedOursRows = new Set<number>();
  const matchedTheirsRows = new Set<number>();

  const matchedInOurs: Array<{ baseIndex: number; sideIndex: number }> = [];
  const matchedInTheirs: Array<{ baseIndex: number; sideIndex: number }> = [];

  const alignedBase: AlignedRow[] = baseRows.map((baseRow) => {
    const key = baseRow.key ?? null;
    if (!key) {
      return {
        base: baseRow,
        ours: null,
        theirs: null,
        key,
        ambiguousOurs: true,
        ambiguousTheirs: true,
      };
    }

    const baseList = baseByKeyList.get(key) ?? [];
    const oursList = oursByKeyList.get(key) ?? [];
    const theirsList = theirsByKeyList.get(key) ?? [];
    const baseCount = baseList.length;
    const oursCount = oursList.length;
    const theirsCount = theirsList.length;
    const occIndex = baseOcc.get(baseRow.index) ?? 0;

    let ours: RowRecord | null = null;
    let theirs: RowRecord | null = null;
    let ambiguousOurs = false;
    let ambiguousTheirs = false;
    const pickBestMatch = (
      candidates: RowRecord[],
      similarityFn: (a: RowRecord, b: RowRecord) => number,
      threshold: number,
      delta: number,
    ) => {
      if (candidates.length === 0) return null;
      const scored = candidates
        .map((r) => ({ row: r, score: similarityFn(baseRow, r) }))
        .sort((a, b) => b.score - a.score);
      const best = scored[0];
      const second = scored[1];
      if (!best || best.score < threshold) return null;
      if (second && best.score - second.score < delta) return null;
      return best.row;
    };

    if (oursCount === 0) {
      const candidates = oursRows.filter((r) => !matchedOursRows.has(r.index));
      const best = pickBestMatch(candidates, rowSimilarityIgnoringKey, rowSimilarityThreshold, 0.05);
      if (best) ours = best;
      else ours = null;
    } else if (oursCount === 1 && baseCount === 1) {
      ours = oursList[0] ?? null;
    } else if (oursCount === baseCount && baseCount > 0) {
      ours = oursList[occIndex] ?? null;
    } else {
      const candidates = oursList.filter((r) => !matchedOursRows.has(r.index));
      if (candidates.length === 1) {
        const only = candidates[0];
        if (rowSimilarity(baseRow, only) >= rowSimilarityThreshold) ours = only;
        else ambiguousOurs = true;
      } else {
        const best = pickBestMatch(candidates, rowSimilarity, rowSimilarityThreshold, 0.1);
        if (best) ours = best;
        else ambiguousOurs = true;
      }
    }

    if (theirsCount === 0) {
      const candidates = theirsRows.filter((r) => !matchedTheirsRows.has(r.index));
      const best = pickBestMatch(candidates, rowSimilarityIgnoringKey, rowSimilarityThreshold, 0.05);
      if (best) theirs = best;
      else theirs = null;
    } else if (theirsCount === 1 && baseCount === 1) {
      theirs = theirsList[0] ?? null;
    } else if (theirsCount === baseCount && baseCount > 0) {
      theirs = theirsList[occIndex] ?? null;
    } else {
      const candidates = theirsList.filter((r) => !matchedTheirsRows.has(r.index));
      if (candidates.length === 1) {
        const only = candidates[0];
        if (rowSimilarity(baseRow, only) >= rowSimilarityThreshold) theirs = only;
        else ambiguousTheirs = true;
      } else {
        const best = pickBestMatch(candidates, rowSimilarity, rowSimilarityThreshold, 0.1);
        if (best) theirs = best;
        else ambiguousTheirs = true;
      }
    }

    if (ours) {
      matchedOursRows.add(ours.index);
      matchedInOurs.push({ baseIndex: baseRow.index, sideIndex: ours.index });
    }
    if (theirs) {
      matchedTheirsRows.add(theirs.index);
      matchedInTheirs.push({ baseIndex: baseRow.index, sideIndex: theirs.index });
    }

    return {
      base: baseRow,
      ours,
      theirs,
      key,
      ambiguousOurs,
      ambiguousTheirs,
    };
  });

  matchedInOurs.sort((a, b) => a.sideIndex - b.sideIndex);
  matchedInTheirs.sort((a, b) => a.sideIndex - b.sideIndex);

  const gapsOurs = new Map<number, RowRecord[]>();
  const gapsTheirs = new Map<number, RowRecord[]>();

  const pushGap = (gaps: Map<number, RowRecord[]>, gap: number, row: RowRecord) => {
    if (!gaps.has(gap)) gaps.set(gap, []);
    gaps.get(gap)!.push(row);
  };

  const placeInGaps = (
    rows: RowRecord[],
    matchedRowIndices: Set<number>,
    matchedPairs: Array<{ baseIndex: number; sideIndex: number }>,
    gaps: Map<number, RowRecord[]>,
  ) => {
    const matchedBaseBySideIndex = matchedPairs.slice().sort((a, b) => a.sideIndex - b.sideIndex);
    for (const row of rows) {
      if (matchedRowIndices.has(row.index)) continue;
      let gap = -1;
      for (const p of matchedBaseBySideIndex) {
        if (p.sideIndex < row.index) gap = p.baseIndex;
        if (p.sideIndex >= row.index) break;
      }
      pushGap(gaps, gap, row);
    }
  };

  placeInGaps(oursRows, matchedOursRows, matchedInOurs, gapsOurs);
  placeInGaps(theirsRows, matchedTheirsRows, matchedInTheirs, gapsTheirs);

  const aligned: AlignedRow[] = [];
  const addGapRows = (gapIndex: number) => {
    const oursGap = gapsOurs.get(gapIndex) ?? [];
    const theirsGap = gapsTheirs.get(gapIndex) ?? [];
    for (const r of oursGap) {
      const ambiguous = !r.key;
      aligned.push({ ours: r, key: r.key ?? null, ambiguousOurs: ambiguous });
    }
    for (const r of theirsGap) {
      const ambiguous = !r.key;
      aligned.push({ theirs: r, key: r.key ?? null, ambiguousTheirs: ambiguous });
    }
  };

  addGapRows(-1);
  for (const baseRow of alignedBase) {
    aligned.push(baseRow);
    addGapRows(baseRow.base?.index ?? -1);
  }

  return { aligned, ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
};

const alignRowsBySequence = (
  baseRows: RowRecord[],
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  const buildTokens = (rows: RowRecord[]) =>
    rows.map((r) => r.values.map((v) => normalizeCellValue(v)).join('||'));

  const similarityThreshold = 0.7;
  const similarityDelta = 0.05;
  const windowSize = 3;

  const alignOneSide = (sideRows: RowRecord[]) => {
    const baseTokens = buildTokens(baseRows);
    const sideTokens = buildTokens(sideRows);
    const ops = myersDiff(baseTokens, sideTokens);
    const matched = new Map<number, number>();
    const deletes = new Set<number>();
    const inserts = new Set<number>();
    for (const op of ops) {
      const hasBase = (idx: number) => idx >= 0 && idx < baseRows.length;
      const hasSide = (idx: number) => idx >= 0 && idx < sideRows.length;
      if (op.type === 'equal') {
        if (hasBase(op.aIndex) && hasSide(op.bIndex)) {
          matched.set(op.aIndex, op.bIndex);
        }
      } else if (op.type === 'delete') {
        if (hasBase(op.aIndex)) deletes.add(op.aIndex);
      } else {
        if (hasSide(op.bIndex)) inserts.add(op.bIndex);
      }
    }

    const unmatchedDeletes = new Set<number>(deletes);
    const unmatchedInserts = new Set<number>(inserts);

    // 浼樺厛鍖归厤鈥滃畬鍏ㄧ浉鍚屸€濈殑琛岋紙token 鐩稿悓锛夛紝閬垮厤閲嶅琛岄€犳垚閿欓厤
    const insertByToken = new Map<string, number[]>();
    for (const idx of unmatchedInserts) {
      const token = sideTokens[idx] ?? '';
      if (!insertByToken.has(token)) insertByToken.set(token, []);
      insertByToken.get(token)!.push(idx);
    }
    insertByToken.forEach((list) => list.sort((a, b) => a - b));

    const matchExactToken = (baseIndex: number) => {
      const token = baseTokens[baseIndex] ?? '';
      const list = insertByToken.get(token);
      if (!list || list.length === 0) return null;
      // 閫夋嫨璺濈鏈熸湜浣嶇疆鏈€杩戠殑鎻掑叆鐐?
      const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
      matchedPairs.sort((a, b) => a.baseIndex - b.baseIndex);
      const expected = estimateSideIndex(baseIndex, matchedPairs);
      let bestPos = 0;
      let bestDist = Math.abs(list[0] - expected);
      for (let i = 1; i < list.length; i += 1) {
        const dist = Math.abs(list[i] - expected);
        if (dist < bestDist) {
          bestDist = dist;
          bestPos = i;
        }
      }
      const sideIndex = list.splice(bestPos, 1)[0];
      if (list.length === 0) insertByToken.delete(token);
      return sideIndex ?? null;
    };

    for (const baseIndex of deletes) {
      const sideIndex = matchExactToken(baseIndex);
      if (sideIndex == null) continue;
      matched.set(baseIndex, sideIndex);
      unmatchedDeletes.delete(baseIndex);
      unmatchedInserts.delete(sideIndex);
    }

    const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
    matchedPairs.sort((a, b) => a.baseIndex - b.baseIndex);

    const ambiguousBase = new Set<number>();
    const ambiguousSide = new Set<number>();
    for (const baseIndex of unmatchedDeletes) {
      const baseRow = baseRows[baseIndex];
      if (!baseRow) continue;
      const expected = estimateSideIndex(baseIndex, matchedPairs);
      const candidates: Array<{ index: number; score: number }> = [];
      for (const sideIndex of unmatchedInserts) {
        if (sideIndex < expected - windowSize || sideIndex > expected + windowSize) continue;
        const sideRow = sideRows[sideIndex];
        if (!sideRow) continue;
        const score = rowSimilarity(baseRow, sideRow);
        if (score >= similarityThreshold) candidates.push({ index: sideIndex, score });
      }
      if (candidates.length === 0) continue;
      candidates.sort((a, b) => b.score - a.score);
      const best = candidates[0];
      const second = candidates[1];
      if (second && second.score >= similarityThreshold && best.score - second.score < similarityDelta) {
        ambiguousBase.add(baseIndex);
        candidates.forEach((c) => ambiguousSide.add(c.index));
        continue;
      }
      matched.set(baseIndex, best.index);
      unmatchedInserts.delete(best.index);
    }

    return { matched, unmatchedInserts, ambiguousBase, ambiguousSide };
  };

  const oursAlign = alignOneSide(oursRows);
  const theirsAlign = alignOneSide(theirsRows);

  const gapsOurs = new Map<number, RowRecord[]>();
  const gapsTheirs = new Map<number, RowRecord[]>();

  const buildGaps = (
    sideRows: RowRecord[],
    matched: Map<number, number>,
    unmatchedInserts: Set<number>,
    gaps: Map<number, RowRecord[]>,
  ) => {
    const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
    matchedPairs.sort((a, b) => a.sideIndex - b.sideIndex);
    for (const sideIndex of unmatchedInserts) {
      const row = sideRows[sideIndex];
      if (!row) continue;
      let gap = -1;
      for (const p of matchedPairs) {
        if (p.sideIndex < sideIndex) gap = p.baseIndex;
        if (p.sideIndex >= sideIndex) break;
      }
      if (!gaps.has(gap)) gaps.set(gap, []);
      gaps.get(gap)!.push(row);
    }
  };

  buildGaps(oursRows, oursAlign.matched, oursAlign.unmatchedInserts, gapsOurs);
  buildGaps(theirsRows, theirsAlign.matched, theirsAlign.unmatchedInserts, gapsTheirs);

  const aligned: AlignedRow[] = [];
  const addGapRows = (gapIndex: number) => {
    const oursGap = gapsOurs.get(gapIndex) ?? [];
    const theirsGap = gapsTheirs.get(gapIndex) ?? [];
    for (const r of oursGap) {
      aligned.push({ ours: r, ambiguousOurs: oursAlign.ambiguousSide.has(r.index) });
    }
    for (const r of theirsGap) {
      aligned.push({ theirs: r, ambiguousTheirs: theirsAlign.ambiguousSide.has(r.index) });
    }
  };

  addGapRows(-1);
  for (let i = 0; i < baseRows.length; i += 1) {
    const baseRow = baseRows[i];
    const oursIndex = oursAlign.matched.get(i);
    const theirsIndex = theirsAlign.matched.get(i);
    aligned.push({
      base: baseRow,
      ours: typeof oursIndex === 'number' ? oursRows[oursIndex] : null,
      theirs: typeof theirsIndex === 'number' ? theirsRows[theirsIndex] : null,
      ambiguousOurs: oursAlign.ambiguousBase.has(i) || (typeof oursIndex === 'number' && oursAlign.ambiguousSide.has(oursIndex)),
      ambiguousTheirs:
        theirsAlign.ambiguousBase.has(i) || (typeof theirsIndex === 'number' && theirsAlign.ambiguousSide.has(theirsIndex)),
    });
    addGapRows(i);
  }

  return { aligned, ambiguousOurs: oursAlign.ambiguousSide, ambiguousTheirs: theirsAlign.ambiguousSide };
};

// Align rows by content using unique anchors, then diff segments to reduce misalignment noise.
const alignRowsByContent = (
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  if (oursRows.length === 0 && theirsRows.length === 0) {
    return { aligned: [], ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
  }
  if (oursRows.length === 0) {
    return { aligned: theirsRows.map((r) => ({ theirs: r })), ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
  }
  if (theirsRows.length === 0) {
    return {
      aligned: oursRows.map((r) => ({ base: r, ours: r })),
      ambiguousOurs: new Set(),
      ambiguousTheirs: new Set(),
    };
  }

  const tokenOf = (r: RowRecord) => r.values.map((v) => normalizeCellValue(v)).join('||');
  const oursTokens = oursRows.map((r) => tokenOf(r));
  const theirsTokens = theirsRows.map((r) => tokenOf(r));

  const countTokens = (tokens: string[]) => {
    const m = new Map<string, number>();
    tokens.forEach((t) => m.set(t, (m.get(t) ?? 0) + 1));
    return m;
  };
  const oursCount = countTokens(oursTokens);
  const theirsCount = countTokens(theirsTokens);
  const theirsUniqueIndex = new Map<string, number>();
  theirsTokens.forEach((t, idx) => {
    if ((theirsCount.get(t) ?? 0) === 1) theirsUniqueIndex.set(t, idx);
  });

  const anchors: Array<{ o: number; t: number }> = [];
  oursTokens.forEach((t, o) => {
    if ((oursCount.get(t) ?? 0) !== 1) return;
    const tIdx = theirsUniqueIndex.get(t);
    if (typeof tIdx === 'number') anchors.push({ o, t: tIdx });
  });

  const selectIncreasingAnchors = (pairs: Array<{ o: number; t: number }>) => {
    if (pairs.length === 0) return [];
    // pairs are already in ours order; compute LIS on t
    const tails: number[] = [];
    const prev = new Array(pairs.length).fill(-1);
    for (let i = 0; i < pairs.length; i += 1) {
      const tVal = pairs[i].t;
      let l = 0;
      let r = tails.length;
      while (l < r) {
        const m = Math.floor((l + r) / 2);
        if (pairs[tails[m]].t < tVal) l = m + 1;
        else r = m;
      }
      if (l > 0) prev[i] = tails[l - 1];
      if (l === tails.length) tails.push(i);
      else tails[l] = i;
    }
    const result: Array<{ o: number; t: number }> = [];
    let k = tails[tails.length - 1];
    while (k >= 0) {
      result.push(pairs[k]);
      k = prev[k];
    }
    return result.reverse();
  };

  const inOrderAnchors = selectIncreasingAnchors(anchors);
  if (inOrderAnchors.length === 0) {
    // fallback to sequence alignment with ours as base
    return alignRowsBySequence(oursRows, oursRows, theirsRows);
  }

  const aligned: AlignedRow[] = [];
  const addSegment = (oStart: number, oEnd: number, tStart: number, tEnd: number) => {
    const oSeg = oursRows.slice(oStart, oEnd);
    const tSeg = theirsRows.slice(tStart, tEnd);
    if (oSeg.length === 0 && tSeg.length === 0) return;
    if (oSeg.length === 0) {
      tSeg.forEach((r) => aligned.push({ theirs: r }));
      return;
    }
    if (tSeg.length === 0) {
      oSeg.forEach((r) => aligned.push({ base: r, ours: r }));
      return;
    }
    const segAligned = alignRowsBySequence(oSeg, oSeg, tSeg).aligned;
    aligned.push(...segAligned);
  };

  let prevO = -1;
  let prevT = -1;
  for (const anchor of inOrderAnchors) {
    addSegment(prevO + 1, anchor.o, prevT + 1, anchor.t);
    aligned.push({
      base: oursRows[anchor.o],
      ours: oursRows[anchor.o],
      theirs: theirsRows[anchor.t],
    });
    prevO = anchor.o;
    prevT = anchor.t;
  }
  addSegment(prevO + 1, oursRows.length, prevT + 1, theirsRows.length);

  return { aligned, ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
};

const buildMergeSheetWithRowAlign = (
  baseWs: any,
  oursWs: any,
  theirsWs: any,
  primaryKeyCol: number,
  frozenRowCount: number,
  rowSimilarityThreshold: number,
): MergeSheetData => {
  const sheetsEqualByCoordinate = (a: any, b: any) => {
    const maxRow = Math.max(getRowCount(a), getRowCount(b));
    const maxCol = Math.max(getColCount(a), getColCount(b));
    for (let r = 1; r <= maxRow; r += 1) {
      const rowA = a.getRow(r);
      const rowB = b.getRow(r);
      for (let c = 1; c <= maxCol; c += 1) {
        const av = normalizeCellValue(getSimpleValueForMerge(rowA.getCell(c)?.value));
        const bv = normalizeCellValue(getSimpleValueForMerge(rowB.getCell(c)?.value));
        if (av !== bv) return false;
      }
    }
    return true;
  };
  const getRowCount = (ws: any) =>
    (ws?.actualRowCount ?? 0) > 0 ? ws.actualRowCount : ws?.rowCount ?? 0;
  const getColCount = (ws: any) =>
    (ws?.actualColumnCount ?? 0) > 0 ? ws.actualColumnCount : ws?.columnCount ?? 0;
  // note: hasExactDiff will be derived from visible diff cells (ours/theirs/conflict)
  const detectKeyColByThreshold = (
    rows: RowRecord[],
    totalCols: number,
    minCoverage: number,
    minUniq: number,
  ) => {
    const total = rows.length;
    if (total === 0) return null;
    const minNonEmpty = Math.max(3, Math.floor(total * minCoverage));
    let bestCol: number | null = null;
    let bestScore = 0;
    for (let col = 1; col <= totalCols; col += 1) {
      let nonEmpty = 0;
      const uniq = new Set<string>();
      for (const row of rows) {
        const v = normalizeKeyValue(row.values[col - 1] ?? null);
        if (v == null) continue;
        nonEmpty += 1;
        uniq.add(v);
      }
      if (nonEmpty < minNonEmpty) continue;
      const coverage = nonEmpty / total;
      const uniqueness = uniq.size / Math.max(1, nonEmpty);
      if (coverage < minCoverage || uniqueness < minUniq) continue;
      const score = coverage * uniqueness;
      if (score > bestScore) {
        bestScore = score;
        bestCol = col;
      }
    }
    return bestCol;
  };
  const detectImplicitKeyCol = (rows: RowRecord[], totalCols: number) =>
    detectKeyColByThreshold(rows, totalCols, 0.8, 0.9);
  const detectWeakKeyCol = (rows: RowRecord[], totalCols: number) =>
    detectKeyColByThreshold(rows, totalCols, 0.6, 0.9);
  const detectHeaderKeyCol = (ws: any, totalCols: number, headerRows: number) => {
    const maxHeader = Math.max(1, Math.min(Math.floor(headerRows), 3));
    for (let r = 1; r <= maxHeader; r += 1) {
      const row = ws.getRow(r);
      for (let c = 1; c <= totalCols; c += 1) {
        const raw = getSimpleValueForMerge(row.getCell(c)?.value);
        if (raw == null) continue;
        const text = String(raw).trim();
        if (!text) continue;
        if (/id/i.test(text) || /缂栧彿|涓婚敭/.test(text)) {
          return c;
        }
      }
    }
    return null;
  };
  const applyKeyFromColumn = (rows: RowRecord[], col: number): RowRecord[] =>
    rows.map((r) => ({
      ...r,
      key: col >= 1 ? normalizeKeyValue(r.values[col - 1] ?? null) : null,
    }));
  const rawColCount = Math.max(
    baseWs?.actualColumnCount ?? baseWs?.columnCount ?? 0,
    oursWs?.actualColumnCount ?? oursWs?.columnCount ?? 0,
    theirsWs?.actualColumnCount ?? theirsWs?.columnCount ?? 0,
  );
  const headerCount = Math.max(0, Math.floor(frozenRowCount));
  const baseWsForAlign = IGNORE_BASE_IN_DIFF ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, headerCount);
  const colCount = Math.max(alignedColumns.length, 0);
  const useKey = primaryKeyCol >= 1 && primaryKeyCol <= rawColCount;
  if (IGNORE_BASE_IN_DIFF && sheetsEqualByCoordinate(oursWs, theirsWs)) {
    return { sheetName: baseWs.name, cells: [], rowsMeta: [], hasExactDiff: false };
  }
  const mapRawToAligned = (rawCol: number, side: 'base' | 'ours' | 'theirs'): number | null => {
    if (rawCol < 1) return null;
    const idx = alignedColumns.findIndex((c) =>
      side === 'base' ? c.baseCol === rawCol : side === 'ours' ? c.oursCol === rawCol : c.theirsCol === rawCol,
    );
    return idx >= 0 ? idx + 1 : null;
  };
  const keyColAligned = useKey ? mapRawToAligned(primaryKeyCol, 'ours') ?? -1 : -1;

  const baseRows = buildRowRecordsAligned(baseWsForAlign, alignedColumns, keyColAligned, 'base').filter(
    (r) => r.rowNumber > headerCount,
  );
  const oursRows = buildRowRecordsAligned(oursWs, alignedColumns, keyColAligned, 'ours').filter(
    (r) => r.rowNumber > headerCount,
  );
  const theirsRows = buildRowRecordsAligned(theirsWs, alignedColumns, keyColAligned, 'theirs').filter(
    (r) => r.rowNumber > headerCount,
  );
  const implicitKeyCol = useKey ? null : detectImplicitKeyCol(baseRows, colCount);
  const headerKeyColRaw =
    !useKey && implicitKeyCol == null ? detectHeaderKeyCol(baseWsForAlign, rawColCount, headerCount) : null;
  const headerKeyCol = headerKeyColRaw ? mapRawToAligned(headerKeyColRaw, 'base') : null;
  const weakKeyCol =
    !useKey && implicitKeyCol == null && headerKeyCol == null ? detectWeakKeyCol(baseRows, colCount) : null;
  const alignKeyCol = useKey ? keyColAligned ?? -1 : implicitKeyCol ?? headerKeyCol ?? weakKeyCol ?? -1;
  const alignedResult =
    alignKeyCol >= 1
      ? alignRowsByKey(
          applyKeyFromColumn(baseRows, alignKeyCol),
          applyKeyFromColumn(oursRows, alignKeyCol),
          applyKeyFromColumn(theirsRows, alignKeyCol),
          alignKeyCol,
          rowSimilarityThreshold,
        )
      : IGNORE_BASE_IN_DIFF
        ? alignRowsByContent(oursRows, theirsRows)
        : alignRowsBySequence(baseRows, oursRows, theirsRows);

  const aligned = alignedResult.aligned;

  const rowsMeta: MergeRowMeta[] = [];
  // 1) Header rows: compare by fixed row number (no alignment)
  const metaKeyCol = alignKeyCol >= 1 ? alignKeyCol : keyColAligned;
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
    const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
    const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
    const oursSim = rowSimilarity(baseRow, oursRow);
    const theirsSim = rowSimilarity(baseRow, theirsRow);
    rowsMeta.push({
      visualRowNumber: r,
      key: baseRow.key ?? oursRow.key ?? theirsRow.key ?? null,
      baseRowNumber: r,
      oursRowNumber: r,
      theirsRowNumber: r,
      oursSimilarity: oursSim,
      theirsSimilarity: theirsSim,
      oursStatus: computeRowStatus(baseRow, oursRow, false),
      theirsStatus: computeRowStatus(baseRow, theirsRow, false),
    });
  }
  // 2) Body rows: aligned
  aligned.forEach((row, idx) => {
    const visualRowNumber = headerCount + idx + 1;
    const oursSim = row.base && row.ours ? rowSimilarity(row.base, row.ours) : null;
    const theirsSim = row.base && row.theirs ? rowSimilarity(row.base, row.theirs) : null;
    rowsMeta.push({
      visualRowNumber,
      key: alignKeyCol >= 1 ? row.key ?? row.base?.key ?? row.ours?.key ?? row.theirs?.key ?? null : null,
      baseRowNumber: row.base?.rowNumber ?? null,
      oursRowNumber: row.ours?.rowNumber ?? null,
      theirsRowNumber: row.theirs?.rowNumber ?? null,
      oursSimilarity: oursSim,
      theirsSimilarity: theirsSim,
      oursStatus: computeRowStatus(row.base ?? null, row.ours ?? null, row.ambiguousOurs),
      theirsStatus: computeRowStatus(row.base ?? null, row.theirs ?? null, row.ambiguousTheirs),
    });
  });

  const same = (a: SimpleCellValue, b: SimpleCellValue) => normalizeCellValue(a) === normalizeCellValue(b);
  const cells: MergeCell[] = [];
  let hasExactDiff = false;

  // Header rows diff by fixed row number (compare ours vs theirs only)
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
    const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
    const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
    const cols = new Set<number>();
    baseRow.nonEmptyCols.forEach((c) => cols.add(c));
    oursRow.nonEmptyCols.forEach((c) => cols.add(c));
    theirsRow.nonEmptyCols.forEach((c) => cols.add(c));
    for (const col of cols) {
      const baseValue = baseRow.values[col - 1] ?? null;
      const oursValue = oursRow.values[col - 1] ?? null;
      const theirsValue = theirsRow.values[col - 1] ?? null;

      const equalOT = same(oursValue, theirsValue);

      let status: MergeCell['status'];
      let mergedValue: SimpleCellValue = oursValue;

      if (equalOT) {
        status = 'unchanged';
        mergedValue = oursValue;
      } else {
        status = 'conflict';
        mergedValue = oursValue;
      }

      if (status !== 'unchanged') {
        const colMeta = alignedColumns[col - 1];
        cells.push({
          address: makeAddress(col, r),
          row: r,
          col,
          baseCol: colMeta?.baseCol ?? null,
          oursCol: colMeta?.oursCol ?? null,
          theirsCol: colMeta?.theirsCol ?? null,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
        hasExactDiff = true;
      }
    }
  }

  // Body rows diff via alignment (compare ours vs theirs only)
  aligned.forEach((row, visualIndex) => {
    const visualRowNumber = headerCount + visualIndex + 1;
    const cols = new Set<number>();
    row.base?.nonEmptyCols.forEach((c) => cols.add(c));
    row.ours?.nonEmptyCols.forEach((c) => cols.add(c));
    row.theirs?.nonEmptyCols.forEach((c) => cols.add(c));
    if (cols.size === 0) return;

    for (const col of cols) {
      const baseValue = row.base?.values[col - 1] ?? null;
      const oursValue = row.ours?.values[col - 1] ?? null;
      const theirsValue = row.theirs?.values[col - 1] ?? null;

      const equalOT = same(oursValue, theirsValue);

      let status: MergeCell['status'];
      let mergedValue: SimpleCellValue = oursValue;

      if (equalOT) {
        status = 'unchanged';
        mergedValue = oursValue;
      } else {
        status = 'conflict';
        mergedValue = oursValue;
      }

      if (status !== 'unchanged') {
        const addressRow =
          row.ours?.rowNumber ?? row.base?.rowNumber ?? row.theirs?.rowNumber ?? visualRowNumber;
        const colMeta = alignedColumns[col - 1];
        cells.push({
          address: makeAddress(col, addressRow),
          row: visualRowNumber,
          col,
          baseCol: colMeta?.baseCol ?? null,
          oursCol: colMeta?.oursCol ?? null,
          theirsCol: colMeta?.theirsCol ?? null,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
        hasExactDiff = true;
      }
    }
  });

  // 濡傛灉鏈夊樊寮傚垪锛屼负鍐荤粨琛岃ˉ榻愯繖浜涘垪鐨勫唴瀹癸紙鍗充娇鏈彉鍖栵級锛岀敤浜庢樉绀鸿〃澶?鍐荤粨琛屼笂涓嬫枃
  if (headerCount > 0 && cells.length > 0) {
    const diffColumns = new Set<number>(cells.map((c) => c.col));
    if (diffColumns.size > 0) {
      const existing = new Set<string>(cells.map((c) => `${c.row}:${c.col}`));
      for (let r = 1; r <= headerCount; r += 1) {
        const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
        const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
        const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
        for (const col of diffColumns) {
          const key = `${r}:${col}`;
          if (existing.has(key)) continue;
          const baseValue = baseRow.values[col - 1] ?? null;
          const oursValue = oursRow.values[col - 1] ?? null;
          const theirsValue = theirsRow.values[col - 1] ?? null;
          const colMeta = alignedColumns[col - 1];
          cells.push({
            address: makeAddress(col, r),
            row: r,
            col,
            baseCol: colMeta?.baseCol ?? null,
            oursCol: colMeta?.oursCol ?? null,
            theirsCol: colMeta?.theirsCol ?? null,
            baseValue,
            oursValue,
            theirsValue,
            status: 'unchanged',
            mergedValue: baseValue,
          });
          existing.add(key);
        }
      }
    }
  }
  cells.sort((a, b) => a.row - b.row || a.col - b.col);

  return {
    sheetName: baseWs.name,
    cells,
    rowsMeta,
    hasExactDiff,
    columnsMeta: alignedColumns.map((c, idx) => ({
      col: idx + 1,
      baseCol: c.baseCol ?? null,
      oursCol: c.oursCol ?? null,
      theirsCol: c.theirsCol ?? null,
    })),
  };
};

// 绠€鍗曠紦瀛橈細鍚屼竴娆″簲鐢ㄧ敓鍛藉懆鏈熷唴閲嶅璇诲彇鍚屼竴涓?xlsx 鏃跺鐢?workbook锛屽噺灏?IO
const workbookCache = new Map<string, Workbook>();

const loadWorkbookCached = async (filePath: string): Promise<Workbook> => {
  const hit = workbookCache.get(filePath);
  if (hit) return hit;
  const wb = new Workbook();
  await wb.xlsx.readFile(filePath);
  workbookCache.set(filePath, wb);
  return wb;
};

const getWorksheetSafe = (wb: Workbook, sheetName?: string, sheetIndex?: number): any => {
  if (sheetName) {
    const byName = wb.getWorksheet(sheetName);
    if (byName) return byName;
  }
  if (typeof sheetIndex === 'number' && sheetIndex >= 0 && sheetIndex < wb.worksheets.length) {
    return wb.worksheets[sheetIndex];
  }
  return wb.worksheets[0];
};

const buildMergeSheetsForWorkbooks = async (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  primaryKeyCol: number,
  frozenRowCount: number,
  rowSimilarityThreshold: number,
) => {
  // 澶嶇敤缂撳瓨锛岄伩鍏嶆瘡娆¤皟鍙?閲嶆柊 diff 鏃堕噸澶嶄粠纾佺洏璇诲彇
  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(basePath),
    loadWorkbookCached(oursPath),
    loadWorkbookCached(theirsPath),
  ]);

  const baseList = baseWb.worksheets;
  const oursList = oursWb.worksheets;
  const theirsList = theirsWb.worksheets;

  const baseByName = new Map<string, { ws: any; idx: number }>();
  baseList.forEach((ws, idx) => {
    if (!baseByName.has(ws.name)) baseByName.set(ws.name, { ws, idx });
  });
  const oursByName = new Map<string, { ws: any; idx: number }>();
  oursList.forEach((ws, idx) => {
    if (!oursByName.has(ws.name)) oursByName.set(ws.name, { ws, idx });
  });
  const theirsByName = new Map<string, { ws: any; idx: number }>();
  theirsList.forEach((ws, idx) => {
    if (!theirsByName.has(ws.name)) theirsByName.set(ws.name, { ws, idx });
  });

  // 瑙勫垯锛氫紭鍏堟寜鍚屽悕宸ヤ綔琛ㄥ榻愶紱瀵瑰墿浣欐湭鍖归厤鐨勫伐浣滆〃锛屽啀鎸夌储寮曞榻愶紙绗?1 寮犲绗?1 寮犫€︹€︼級銆?
  const usedBaseIdx = new Set<number>();
  const usedOursIdx = new Set<number>();
  const usedTheirsIdx = new Set<number>();

  const mergeSheets: MergeSheetData[] = [];

  // 1) 鍚屽悕鍖归厤锛氫互 base 鐨勯『搴忎负鍑?
  for (let i = 0; i < baseList.length; i += 1) {
    const baseWs = baseList[i];
    const oursHit = oursByName.get(baseWs.name);
    const theirsHit = theirsByName.get(baseWs.name);
    if (!oursHit || !theirsHit) continue;

    usedBaseIdx.add(i);
    usedOursIdx.add(oursHit.idx);
    usedTheirsIdx.add(theirsHit.idx);

    mergeSheets.push(
      buildMergeSheetWithRowAlign(baseWs, oursHit.ws, theirsHit.ws, primaryKeyCol, frozenRowCount, rowSimilarityThreshold),
    );
  }

  // 2) 绱㈠紩鍏滃簳锛氫粎瀵光€滃悓涓€ idx 鍦ㄤ笁杈归兘娌¤鐢ㄨ繃鈥濈殑浣嶇疆鍋氬榻?
  const count = Math.min(baseList.length, oursList.length, theirsList.length);
  for (let idx = 0; idx < count; idx += 1) {
    if (usedBaseIdx.has(idx) || usedOursIdx.has(idx) || usedTheirsIdx.has(idx)) continue;
    usedBaseIdx.add(idx);
    usedOursIdx.add(idx);
    usedTheirsIdx.add(idx);
    mergeSheets.push(
      buildMergeSheetWithRowAlign(baseList[idx], oursList[idx], theirsList[idx], primaryKeyCol, frozenRowCount, rowSimilarityThreshold),
    );
  }

  return { basePath, oursPath, theirsPath, mergeSheets };
};

const normalizeThreeWayResult = (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  mergeSheets: MergeSheetData[],
) => {
  const emptySheet: MergeSheetData = { sheetName: '', cells: [], rowsMeta: [] };
  return {
    basePath,
    oursPath,
    theirsPath,
    sheet: mergeSheets[0] ?? emptySheet,
    sheets: mergeSheets,
  };
};

// IPC types
interface SheetCell {
  address: string; // e.g. "A1"
  row: number;
  col: number;
  value: string | number | null;
}

type RowStatus = 'unchanged' | 'added' | 'deleted' | 'modified' | 'ambiguous';

interface MergeRowMeta {
  /** 瑙嗚琛屽彿锛坉iff/merge 瑙嗗浘涓殑 1-based 琛屽彿锛?*/
  visualRowNumber: number;
  /** 濡傛灉鍚敤浜嗕富閿垪锛岃繖閲岃褰曚富閿紙normalize 鍚庯級 */
  key?: string | null;
  /** 涓夋柟鏂囦欢涓悇鑷搴旂殑鍘熷琛屽彿锛?-based锛夛紱涓嶅瓨鍦ㄥ垯涓?null */
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  /** 琛岀浉浼煎害锛堢浉瀵?base锛岃寖鍥?0-1锛?*/
  oursSimilarity?: number | null;
  theirsSimilarity?: number | null;
  /** 璇ヨ瑙夎鍦ㄥ搴?side 鐩稿 base 鐨勭姸鎬?*/
  oursStatus: RowStatus;
  theirsStatus: RowStatus;
}

interface SheetData {
  sheetName: string;
  rows: SheetCell[][];
}

interface MergeCell {
  address: string;
  row: number;
  col: number;
  baseCol?: number | null;
  oursCol?: number | null;
  theirsCol?: number | null;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;
}
interface MergeColumnMeta {
  col: number; // aligned column index (1-based)
  baseCol: number | null;
  oursCol: number | null;
  theirsCol: number | null;
}

interface MergeSheetData {
  sheetName: string;
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  hasExactDiff?: boolean;
  columnsMeta?: MergeColumnMeta[];
}

interface SaveMergeCellInput {
  address: string;
  value: string | number | null;
}
interface SaveMergeRowOp {
  sheetName: string;
  action: 'insert' | 'delete';
  targetRowNumber: number; // 1-based in template (ours)
  values?: (string | number | null)[];
  visualRowNumber?: number;
}

interface SaveMergeColOp {
  sheetName: string;
  action: 'insert' | 'delete';
  targetColNumber: number; // 1-based in template (ours)
  alignedColNumber?: number; // 1-based aligned column index
  values?: (string | number | null)[];
  source?: 'theirs' | 'base' | 'ours';
}

interface SaveMergeRequest {
  templatePath: string;
  cells: SaveMergeCellInput[];
  rowOps?: SaveMergeRowOp[];
  colOps?: SaveMergeColOp[];
  basePath?: string;
  oursPath?: string;
  theirsPath?: string;
}

interface SaveMergeResponse {
  success: boolean;
  filePath?: string;
  cancelled?: boolean;
  errorMessage?: string;
}

let currentFilePath: string | null = null;

/**
 * 澶勭悊娓叉煋杩涚▼璇锋眰锛氶€夋嫨骞舵墦寮€涓€涓?Excel 鏂囦欢銆?
 *
 * 杩斿洖锛氭枃浠惰矾寰?+ 鎵€鏈夊伐浣滆〃鐨勪簩缁村崟鍏冩牸鏁版嵁锛堜粎鍖呭惈鈥滃€尖€濓級锛?
 * 鐢ㄤ簬鍗曟枃浠舵煡鐪?缂栬緫妯″紡銆?
 */
ipcMain.handle('excel:open', async () => {
  if (!mainWindow) return null;

  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });

  if (canceled || filePaths.length === 0) {
    return null;
  }

  const filePath = filePaths[0];
  currentFilePath = filePath;

  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);

  const buildSheetData = (worksheet: Worksheet): SheetData => {
    const rows: SheetCell[][] = [];

    const getSimpleValue = (raw: CellValue): string | number | null => {
      if (raw === null || raw === undefined) return null;

      // Date
      if (raw instanceof Date) {
        // 淇濇寔鍙鎬э紝閬垮厤鏄剧ず涓?[object Object]
        return raw.toISOString();
      }

      // 瀵屾枃鏈細raw.richText 鏄竴涓寘鍚?{ text } 鐨勬暟缁?
      if (typeof raw === 'object' && Array.isArray((raw as any).richText)) {
        const parts = (raw as any).richText
          .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
          .join('');
        return parts;
      }

      // Hyperlink / text-like objects
      if (typeof raw === 'object' && raw && 'text' in (raw as any)) {
        const t = (raw as any).text;
        if (t === null || t === undefined) return null;
        return typeof t === 'string' || typeof t === 'number' ? (t as any) : String(t);
      }

      // Formula / shared formula 绛夛細浼樺厛鏄剧ず result
      if (typeof raw === 'object' && raw && 'result' in (raw as any)) {
        const r = (raw as any).result;
        if (r === null || r === undefined) return null;
        if (typeof r === 'string' || typeof r === 'number') return r;
        if (r instanceof Date) return r.toISOString();
        return String(r);
      }

      if (typeof raw === 'string' || typeof raw === 'number') {
        return raw;
      }

      // 鍏滃簳锛氬敖閲?JSON 搴忓垪鍖栵紝閬垮厤 [object Object]
      if (typeof raw === 'object') {
        try {
          return JSON.stringify(raw);
        } catch {
          return String(raw);
        }
      }

      return String(raw);
    };

    // 閲嶈锛氱‘淇濇瘡涓€琛岀殑鍒楁暟涓€鑷淬€?
    // 鍚﹀垯浼氬嚭鐜扳€滄暟鎹鍒楁暟 > 琛ㄥご/鍐荤粨琛屽垪鏁扳€濋€犳垚閿欎綅銆?
    const maxRow =
      (worksheet as any).actualRowCount && (worksheet as any).actualRowCount > 0
        ? (worksheet as any).actualRowCount
        : worksheet.rowCount;
    const maxCol =
      (worksheet as any).actualColumnCount && (worksheet as any).actualColumnCount > 0
        ? (worksheet as any).actualColumnCount
        : worksheet.columnCount;

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      const rowCells: SheetCell[] = [];
      const row = worksheet.getRow(rowNumber);
      for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
        const cell = row.getCell(colNumber);
        const value = getSimpleValue(cell.value as any);
        rowCells.push({
          address: cell.address,
          row: rowNumber,
          col: colNumber,
          value,
        });
      }
      rows.push(rowCells);
    }

    return {
      sheetName: worksheet.name,
      rows,
    };
  };

  const sheets: SheetData[] = workbook.worksheets.map((ws) => buildSheetData(ws));

  return { filePath, sheet: sheets[0], sheets };
});

interface CellChange {
  address: string;
  newValue: string | number | null;
}
interface GetSheetDataRequest {
  path: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
}

/**
 * 灏嗗崟鏂囦欢缂栬緫妯″紡涓嬬敤鎴蜂慨鏀硅繃鐨勫崟鍏冩牸鍐欏洖鍘熷 Excel 鏂囦欢銆?
 *
 * 鍙慨鏀瑰崟鍏冩牸鐨?value锛屼笉鍔ㄦ牱寮?鍏紡绛夋牸寮忎俊鎭€?
 */
ipcMain.handle('excel:saveChanges', async (_event, req: CellChange[] | { changes: CellChange[]; sheetName?: string; sheetIndex?: number }) => {
  if (!currentFilePath) {
    throw new Error('No Excel file is currently loaded');
  }
  const changes: CellChange[] = Array.isArray(req) ? req : (req?.changes ?? []);
  const sheetName = !Array.isArray(req) ? req?.sheetName : undefined;
  const sheetIndex = !Array.isArray(req) ? req?.sheetIndex : undefined;

  const workbook = new Workbook();
  await workbook.xlsx.readFile(currentFilePath);
  let worksheet = sheetName ? workbook.getWorksheet(sheetName) ?? undefined : undefined;
  if (!worksheet && typeof sheetIndex === 'number' && sheetIndex >= 0) {
    worksheet = workbook.worksheets[sheetIndex];
  }
  if (!worksheet) worksheet = workbook.worksheets[0];

  for (const change of changes) {
    const cell = worksheet.getCell(change.address);
    cell.value = change.newValue as any; // only change value, keep formatting/styles
  }

  await workbook.xlsx.writeFile(currentFilePath);
  // invalidate cache to avoid stale reads
  if (workbookCache.has(currentFilePath)) {
    workbookCache.delete(currentFilePath);
  }

  return { success: true };
});

// 璇诲彇鎸囧畾鏂囦欢鐨勬寚瀹氬伐浣滆〃锛堢敤浜?merge 妯″紡涓嬫樉绀哄叏琛級
ipcMain.handle('excel:getSheetData', async (_event, req: GetSheetDataRequest): Promise<SheetData | null> => {
  if (!req || !req.path) return null;
  const wb = await loadWorkbookCached(req.path);
  const ws = getWorksheetSafe(wb, req.sheetName, req.sheetIndex);
  if (!ws) return null;

  const maxRow =
    (ws as any).actualRowCount && (ws as any).actualRowCount > 0
      ? (ws as any).actualRowCount
      : ws.rowCount;
  const maxCol =
    (ws as any).actualColumnCount && (ws as any).actualColumnCount > 0
      ? (ws as any).actualColumnCount
      : ws.columnCount;

  const rows: SheetCell[][] = [];
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const rowCells: SheetCell[] = [];
    const row = ws.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
      const cell = row.getCell(colNumber);
      const value = getSimpleValueForMerge(cell?.value);
      rowCells.push({
        address: cell.address,
        row: rowNumber,
        col: colNumber,
        value,
      });
    }
    rows.push(rowCells);
  }

  return { sheetName: ws.name, rows };
});

// 淇濆瓨涓夋柟 merge 缁撴灉鍒版柊鐨?Excel 鏂囦欢锛屼粎淇敼鍊硷紝涓嶆敼鏍煎紡
//
// 鍦?git/Fork merge 妯″紡涓嬶細
//   - 濡傛灉鎻愪緵浜?MERGED 鍙傛暟锛屽垯缁撴灉鍐欏洖 MERGED锛?
//   - 鍚﹀垯鍥為€€鍒拌鐩?ours锛?
// 鍦?diff 妯″紡涓嬶細
//   - 鐩存帴瑕嗙洊 ours锛圠OCAL锛夈€?
// 浜や簰寮忔ā寮忎笅锛?
//   - 寮瑰嚭淇濆瓨瀵硅瘽妗嗭紝鐢辩敤鎴烽€夋嫨鐩爣璺緞銆?
ipcMain.handle('excel:saveMergeResult', async (_event, req: SaveMergeRequest): Promise<SaveMergeResponse> => {
  if (!mainWindow) {
    throw new Error('Main window is not available');
  }

  try {
    const { templatePath, cells, rowOps, colOps } = req as {
      templatePath: string;
      cells: { sheetName: string; address: string; value: string | number | null }[];
      rowOps?: SaveMergeRowOp[];
      colOps?: SaveMergeColOp[];
    };
    let targetPath: string | undefined;

    if (cliThreeWayArgs && cliThreeWayArgs.mode === 'merge') {
      // git / Fork merge 妯″紡锛氫紭鍏堝啓鍏?MERGED锛堝伐浣滃尯瀵瑰簲鏂囦欢锛夛紝濡傛灉鍛戒护鍙紶浜?base/ours/theirs 涓変釜鍙傛暟锛屽垯鍥為€€瑕嗙洊 ours銆?
      const oursPath = cliThreeWayArgs.oursPath;
      const mergedPath = cliThreeWayArgs.mergedPath;
      targetPath = mergedPath || oursPath;
    } else if (cliThreeWayArgs && cliThreeWayArgs.mode === 'diff') {
      targetPath = cliThreeWayArgs.oursPath;
    } else {
      const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
        title: '淇濆瓨鍚堝苟鍚庣殑 Excel',
        defaultPath: templatePath,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
      });

      if (canceled || !filePath) {
        return { success: false, cancelled: true };
      }
      targetPath = filePath;
    }

    const workbook = new Workbook();
    await workbook.xlsx.readFile(templatePath);

    // IMPORTANT: 蹇呴』鍏堟墽琛屽垪/琛屾搷浣滐紝鍐嶄慨鏀瑰崟鍏冩牸
    // 鍥犱负鍒?琛屾搷浣滀細鏀瑰彉绱㈠紩锛屽鏋滃厛淇敼鍗曞厓鏍硷紝鍦板潃浼氶敊涔?
    
    const colOpsBySheet = new Map<string, SaveMergeColOp[]>();
    const rowOpsBySheet = new Map<string, SaveMergeRowOp[]>();

    // 1. 鍏堟墽琛屽垪鎿嶄綔
    if (colOps && colOps.length > 0) {
      colOps.forEach((op) => {
        const key = op.sheetName || '';
        if (!colOpsBySheet.has(key)) colOpsBySheet.set(key, []);
        colOpsBySheet.get(key)!.push(op);
      });
      colOpsBySheet.forEach((ops, sheetName) => {
        const ws = workbook.getWorksheet(sheetName) ?? workbook.worksheets[0];
        const sorted = ops.slice().sort((a, b) => {
          const va = a.alignedColNumber ?? 0;
          const vb = b.alignedColNumber ?? 0;
          if (va !== vb) return va - vb;
          return a.targetColNumber - b.targetColNumber;
        });
        // Process deletes first (sorted by col descending to maintain positions)
        const deletes = sorted.filter(op => op.action === 'delete').sort((a, b) => b.targetColNumber - a.targetColNumber);
        for (const op of deletes) {
          const colNumber = Math.max(1, Math.floor(op.targetColNumber));
          if (typeof (ws as any).spliceColumns === 'function') {
            (ws as any).spliceColumns(colNumber, 1);
          } else {
            // fallback: manual delete by shifting cells left
            const maxRow = ws?.actualRowCount ?? ws?.rowCount ?? 0;
            const maxCol = ws?.actualColumnCount ?? ws?.columnCount ?? 0;
            for (let r = 1; r <= maxRow; r += 1) {
              for (let c = colNumber; c < maxCol; c += 1) {
                const from = ws.getRow(r).getCell(c + 1);
                const to = ws.getRow(r).getCell(c);
                to.value = from.value as any;
              }
              // Clear last column
              ws.getRow(r).getCell(maxCol).value = null;
            }
          }
        }
        // Then process inserts (sorted by aligned col ascending)
        const inserts = sorted.filter(op => op.action === 'insert');
        let offset = 0;
        for (const op of inserts) {
          // 修正：targetColNumber 基于原始 ours 列号，需减去前面已执行的 delete 偏移
          let baseCol = Math.max(1, Math.floor(op.targetColNumber));
          for (const delOp of deletes) {
            const delCol = Math.max(1, Math.floor(delOp.targetColNumber));
            if (baseCol > delCol) baseCol -= 1;
          }
          const colNumber = baseCol + offset;
          const maxRow = Math.max(
            ws?.actualRowCount ?? ws?.rowCount ?? 0,
            op.values?.length ?? 0,
          );
          const values: (string | number | null)[] = [];
          for (let i = 0; i < maxRow; i += 1) {
            values.push(op.values && i < op.values.length ? op.values[i] ?? null : null);
          }
          if (typeof (ws as any).spliceColumns === 'function') {
            (ws as any).spliceColumns(colNumber, 0, values);
          } else {
            // fallback: manual insert by shifting cells (rare)
            for (let r = maxRow; r >= 1; r -= 1) {
              for (let c = (ws?.actualColumnCount ?? ws?.columnCount ?? 0); c >= colNumber; c -= 1) {
                const from = ws.getRow(r).getCell(c);
                const to = ws.getRow(r).getCell(c + 1);
                to.value = from.value as any;
              }
              const cell = ws.getRow(r).getCell(colNumber);
              cell.value = values[r - 1] ?? null;
            }
          }
          offset += 1;
        }
      });
    }
    // 2. 鍐嶆墽琛岃鎿嶄綔
    if (rowOps && rowOps.length > 0) {
      rowOps.forEach((op) => {
        const key = op.sheetName || '';
        if (!rowOpsBySheet.has(key)) rowOpsBySheet.set(key, []);
        rowOpsBySheet.get(key)!.push(op);
      });
      rowOpsBySheet.forEach((ops, sheetName) => {
        const ws = workbook.getWorksheet(sheetName) ?? workbook.worksheets[0];
        const sorted = ops.slice().sort((a, b) => {
          const va = a.visualRowNumber ?? 0;
          const vb = b.visualRowNumber ?? 0;
          if (va !== vb) return va - vb;
          return a.targetRowNumber - b.targetRowNumber;
        });
        let offset = 0;
        for (const op of sorted) {
          const baseRow = Math.max(1, Math.floor(op.targetRowNumber));
          const rowNumber = baseRow + offset;
          if (op.action === 'insert') {
            const maxCol = Math.max(
              ws?.actualColumnCount ?? ws?.columnCount ?? 0,
              op.values?.length ?? 0,
            );
            const values: (string | number | null)[] = [];
            for (let i = 0; i < maxCol; i += 1) {
              values.push(op.values && i < op.values.length ? op.values[i] ?? null : null);
            }
            ws.spliceRows(rowNumber, 0, values);
            offset += 1;
          } else if (op.action === 'delete') {
            ws.spliceRows(rowNumber, 1);
            offset -= 1;
          }
        }
      });
    }

    const colLabelToNumber = (label: string): number => {
      const s = label.toUpperCase();
      let n = 0;
      for (let i = 0; i < s.length; i += 1) {
        const code = s.charCodeAt(i);
        if (code < 65 || code > 90) return NaN;
        n = n * 26 + (code - 64);
      }
      return n;
    };
    const parseAddress = (address: string): { col: number; row: number } | null => {
      const m = /^([A-Z]+)(\d+)$/i.exec(address);
      if (!m) return null;
      const col = colLabelToNumber(m[1]);
      const row = Number(m[2]);
      if (!Number.isFinite(col) || !Number.isFinite(row)) return null;
      return { col, row };
    };
    const buildRowMapper = (ops: SaveMergeRowOp[]) => {
      const sorted = ops.slice().sort((a, b) => {
        const va = a.visualRowNumber ?? 0;
        const vb = b.visualRowNumber ?? 0;
        if (va !== vb) return va - vb;
        return a.targetRowNumber - b.targetRowNumber;
      });
      return (row: number): number | null => {
        let r = row;
        let offset = 0;
        for (const op of sorted) {
          const baseRow = Math.max(1, Math.floor(op.targetRowNumber));
          const rowNumber = baseRow + offset;
          if (op.action === 'insert') {
            if (r >= rowNumber) r += 1;
            offset += 1;
          } else {
            if (r === rowNumber) return null;
            if (r > rowNumber) r -= 1;
            offset -= 1;
          }
        }
        return r;
      };
    };
    const buildColMapper = (ops: SaveMergeColOp[]) => {
      const sorted = ops.slice().sort((a, b) => {
        const va = a.alignedColNumber ?? 0;
        const vb = b.alignedColNumber ?? 0;
        if (va !== vb) return va - vb;
        return a.targetColNumber - b.targetColNumber;
      });
      const deletes = sorted
        .filter((op) => op.action === 'delete')
        .sort((a, b) => b.targetColNumber - a.targetColNumber);
      const inserts = sorted.filter((op) => op.action === 'insert');
      return (col: number): number | null => {
        let c = col;
        for (const op of deletes) {
          const colNumber = Math.max(1, Math.floor(op.targetColNumber));
          if (c === colNumber) return null;
          if (c > colNumber) c -= 1;
        }
        let offset = 0;
        for (const op of inserts) {
          // 修正：insert 的 targetColNumber 需按已执行 delete 偏移修正
          let adjustedBase = Math.max(1, Math.floor(op.targetColNumber));
          for (const delOp of deletes) {
            const delCol = Math.max(1, Math.floor(delOp.targetColNumber));
            if (adjustedBase > delCol) adjustedBase -= 1;
          }
          const insertAt = adjustedBase + offset;
          if (c >= insertAt) c += 1;
          offset += 1;
        }
        return c;
      };
    };
    const rowMapperCache = new Map<string, (row: number) => number | null>();
    const colMapperCache = new Map<string, (col: number) => number | null>();
    const getRowMapper = (sheetKey: string) => {
      if (!rowMapperCache.has(sheetKey)) {
        rowMapperCache.set(sheetKey, buildRowMapper(rowOpsBySheet.get(sheetKey) ?? []));
      }
      return rowMapperCache.get(sheetKey)!;
    };
    const getColMapper = (sheetKey: string) => {
      if (!colMapperCache.has(sheetKey)) {
        colMapperCache.set(sheetKey, buildColMapper(colOpsBySheet.get(sheetKey) ?? []));
      }
      return colMapperCache.get(sheetKey)!;
    };

    // 3. 鏈€鍚庝慨鏀瑰崟鍏冩牸鍊硷紙姝ゆ椂鍒?琛岀储寮曞凡缁忕ǔ瀹氾級
    for (const cellInfo of cells) {
      const sheetKey = cellInfo.sheetName || '';
      const ws = workbook.getWorksheet(cellInfo.sheetName) ?? workbook.worksheets[0];
      const parsed = parseAddress(cellInfo.address);
      if (!parsed) continue;
      const newCol = getColMapper(sheetKey)(parsed.col);
      if (newCol == null) continue;
      const newRow = getRowMapper(sheetKey)(parsed.row);
      if (newRow == null) continue;
      const newAddress = makeAddress(newCol, newRow);
      const cell = ws.getCell(newAddress);
      cell.value = cellInfo.value as any;
    }

    normalizeSharedFormulas(workbook);
    await workbook.xlsx.writeFile(targetPath);
    // invalidate cache to avoid stale reads
    if (targetPath && workbookCache.has(targetPath)) {
      workbookCache.delete(targetPath);
    }
    if (templatePath && templatePath !== targetPath && workbookCache.has(templatePath)) {
      workbookCache.delete(templatePath);
    }

    // 濡傛灉鏄€氳繃 git/Fork 鐨?merge 妯″紡鍚姩锛屽苟涓旀湁鏄庣‘鐨勭洰鏍囨枃浠讹紝灏濊瘯鑷姩鎵ц涓€娆?git add
    if (cliThreeWayArgs && cliThreeWayArgs.mode === 'merge' && targetPath) {
      try {
        await gitAddFile(targetPath);
      } catch (e) {
        console.error('git add after merge failed', e);
      }
    }

    return { success: true, filePath: targetPath };
  } catch (err: any) {
    console.error('excel:saveMergeResult failed', err);
    return { success: false, errorMessage: err?.message ?? String(err) };
  }
});

// 涓夋柟 diff锛歜ase / ours / theirs锛屽彧姣旇緝鍗曞厓鏍煎€硷紝蹇界暐鏍煎紡
//
// 杩斿洖缁欐覆鏌撹繘绋嬬殑鏁版嵁鏄細
//   - base / ours / theirs 鐨勬枃浠惰矾寰勶紱
//   - 姣忎釜宸ヤ綔琛ㄧ殑涓夋柟鍗曞厓鏍煎€?+ 宸紓鐘舵€侊紙unchanged / conflict 绛夛級銆?
ipcMain.handle('excel:openThreeWay', async () => {
  if (!mainWindow) return null;
  const primaryKeyCol = 1;
  const frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS;
  const rowSimilarityThreshold = DEFAULT_ROW_SIMILARITY_THRESHOLD;

  if (cliThreeWayArgs) {
    const { basePath, oursPath, theirsPath } = cliThreeWayArgs;
    const { mergeSheets } = await buildMergeSheetsForWorkbooks(
      basePath,
      oursPath,
      theirsPath,
      primaryKeyCol,
      frozenRowCount,
      rowSimilarityThreshold,
    );
    return normalizeThreeWayResult(basePath, oursPath, theirsPath, mergeSheets);
  }

  // 娌℃湁 CLI 鍙傛暟鏃讹紝鍥為€€鍒颁氦浜掑紡閫夋嫨鏂囦欢鐨勬ā寮?
  const pickFile = async (title: string) => {
    const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow!, {
      title,
      filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
      properties: ['openFile'],
    });
    if (canceled || filePaths.length === 0) return null;
    return filePaths[0];
  };

  const basePath = await pickFile('閫夋嫨 base 鐗堟湰 Excel');
  if (!basePath) return null;
  const oursPath = await pickFile('閫夋嫨 ours (褰撳墠鍒嗘敮) Excel');
  if (!oursPath) return null;
  const theirsPath = await pickFile('閫夋嫨 theirs (鍚堝苟鍒嗘敮) Excel');
  if (!theirsPath) return null;

  const { mergeSheets } = await buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    primaryKeyCol,
    frozenRowCount,
    rowSimilarityThreshold,
  );

  return normalizeThreeWayResult(basePath, oursPath, theirsPath, mergeSheets);
});
interface ThreeWayDiffRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  primaryKeyCol: number; // 1-based, -1 means no primary key
  frozenRowCount?: number; // header rows compared by coordinates
  rowSimilarityThreshold?: number; // 0-1
}

ipcMain.handle('excel:computeThreeWayDiff', async (_event, req: ThreeWayDiffRequest) => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const primaryKeyCol =
    typeof req.primaryKeyCol === 'number' && !Number.isNaN(req.primaryKeyCol) ? Math.floor(req.primaryKeyCol) : 1;
  const frozenRowCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const rowSimilarityThreshold =
    typeof req.rowSimilarityThreshold === 'number' && !Number.isNaN(req.rowSimilarityThreshold)
      ? Math.min(1, Math.max(0, req.rowSimilarityThreshold))
      : DEFAULT_ROW_SIMILARITY_THRESHOLD;
  const { mergeSheets } = await buildMergeSheetsForWorkbooks(
    req.basePath,
    req.oursPath,
    req.theirsPath,
    primaryKeyCol,
    frozenRowCount,
    rowSimilarityThreshold,
  );
  return normalizeThreeWayResult(req.basePath, req.oursPath, req.theirsPath, mergeSheets);
});

// 灏?CLI three-way 淇℃伅鏆撮湶缁欐覆鏌撹繘绋嬶紝渚夸簬鑷姩鍔犺浇
ipcMain.handle('excel:getCliThreeWayInfo', async () => {
  if (!cliThreeWayArgs) return null;
  return cliThreeWayArgs;
});

// 璇诲彇涓夋柟鏂囦欢鐨勨€滄煇涓€琛屸€濇暟鎹紝鐢ㄤ簬搴曢儴琛岀骇瀵规瘮瑙嗗浘
interface ThreeWayRowRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
  frozenRowCount?: number;
  rowNumber?: number; // 1-based fallback for all sides
  baseRowNumber?: number | null;
  oursRowNumber?: number | null;
  theirsRowNumber?: number | null;
}

interface ThreeWayRowResult {
  sheetName: string;
  rowNumber?: number;
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  colCount: number;
  base: (string | number | null)[];
  ours: (string | number | null)[];
  theirs: (string | number | null)[];
}
interface ThreeWayRowsRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
  frozenRowCount?: number;
  rows: Array<{
    rowNumber?: number;
    baseRowNumber?: number | null;
    oursRowNumber?: number | null;
    theirsRowNumber?: number | null;
  }>;
}
interface ThreeWayRowsResult {
  sheetName: string;
  colCount: number;
  rows: ThreeWayRowResult[];
}

const normalizeSharedFormulas = (workbook: Workbook) => {
  workbook.worksheets.forEach((ws) => {
    ws.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        const v: any = cell.value as any;
        if (!v || typeof v !== 'object') return;
        const isShared = v.sharedFormula || v.shareType === 'shared';
        if (!isShared) return;
        const model: any = (cell as any).model || {};
        const formula = model.formula || v.formula;
        const result = model.result !== undefined ? model.result : v.result;
        if (formula) {
          cell.value = { formula, result } as any;
          return;
        }
        if (result !== undefined) {
          cell.value = result as any;
          return;
        }
        cell.value = null as any;
      });
    });
  });
};


ipcMain.handle('excel:getThreeWayRow', async (_event, req: ThreeWayRowRequest): Promise<ThreeWayRowResult | null> => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const fallbackRow =
    typeof req.rowNumber === 'number' && !Number.isNaN(req.rowNumber)
      ? Math.max(1, Math.floor(req.rowNumber))
      : null;
  const baseRowNumber =
    typeof req.baseRowNumber === 'number' && !Number.isNaN(req.baseRowNumber)
      ? Math.max(1, Math.floor(req.baseRowNumber))
      : fallbackRow;
  const oursRowNumber =
    typeof req.oursRowNumber === 'number' && !Number.isNaN(req.oursRowNumber)
      ? Math.max(1, Math.floor(req.oursRowNumber))
      : fallbackRow;
  const theirsRowNumber =
    typeof req.theirsRowNumber === 'number' && !Number.isNaN(req.theirsRowNumber)
      ? Math.max(1, Math.floor(req.theirsRowNumber))
      : fallbackRow;

  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(req.basePath),
    loadWorkbookCached(req.oursPath),
    loadWorkbookCached(req.theirsPath),
  ]);

  const baseWs = getWorksheetSafe(baseWb, req.sheetName, req.sheetIndex);
  const oursWs = getWorksheetSafe(oursWb, req.sheetName, req.sheetIndex);
  const theirsWs = getWorksheetSafe(theirsWb, req.sheetName, req.sheetIndex);

  const resolvedSheetName = baseWs?.name ?? req.sheetName ?? '';
  const headerCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const baseWsForAlign = IGNORE_BASE_IN_DIFF ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, headerCount);
  const colCount = alignedColumns.length;

  const readRowAligned = (
    ws: any,
    rowNum: number | null,
    side: 'base' | 'ours' | 'theirs',
  ): (string | number | null)[] => {
    const arr: (string | number | null)[] = [];
    if (!rowNum) {
      for (let col = 1; col <= colCount; col += 1) arr.push(null);
      return arr;
    }
    const row = ws.getRow(rowNum);
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const meta = alignedColumns[i];
      const colNumber =
        side === 'base' ? meta.baseCol : side === 'ours' ? meta.oursCol : meta.theirsCol;
      if (!colNumber) {
        arr.push(null);
        continue;
      }
      const cell = row.getCell(colNumber);
      arr.push(getSimpleValueForMerge(cell?.value));
    }
    return arr;
  };

  return {
    sheetName: resolvedSheetName,
    rowNumber: fallbackRow ?? undefined,
    baseRowNumber: baseRowNumber ?? null,
    oursRowNumber: oursRowNumber ?? null,
    theirsRowNumber: theirsRowNumber ?? null,
    colCount,
    base: readRowAligned(baseWs, baseRowNumber ?? null, 'base'),
    ours: readRowAligned(oursWs, oursRowNumber ?? null, 'ours'),
    theirs: readRowAligned(theirsWs, theirsRowNumber ?? null, 'theirs'),
  };
});
ipcMain.handle('excel:getThreeWayRows', async (_event, req: ThreeWayRowsRequest): Promise<ThreeWayRowsResult | null> => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath || !Array.isArray(req.rows)) return null;

  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(req.basePath),
    loadWorkbookCached(req.oursPath),
    loadWorkbookCached(req.theirsPath),
  ]);

  const baseWs = getWorksheetSafe(baseWb, req.sheetName, req.sheetIndex);
  const oursWs = getWorksheetSafe(oursWb, req.sheetName, req.sheetIndex);
  const theirsWs = getWorksheetSafe(theirsWb, req.sheetName, req.sheetIndex);

  const resolvedSheetName = baseWs?.name ?? req.sheetName ?? '';
  const headerCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const baseWsForAlign = IGNORE_BASE_IN_DIFF ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, headerCount);
  const colCount = alignedColumns.length;

  const readRowAligned = (
    ws: any,
    rowNum: number | null,
    side: 'base' | 'ours' | 'theirs',
  ): (string | number | null)[] => {
    const arr: (string | number | null)[] = [];
    if (!rowNum) {
      for (let col = 1; col <= colCount; col += 1) arr.push(null);
      return arr;
    }
    const row = ws.getRow(rowNum);
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const meta = alignedColumns[i];
      const colNumber =
        side === 'base' ? meta.baseCol : side === 'ours' ? meta.oursCol : meta.theirsCol;
      if (!colNumber) {
        arr.push(null);
        continue;
      }
      const cell = row.getCell(colNumber);
      arr.push(getSimpleValueForMerge(cell?.value));
    }
    return arr;
  };

  const rows: ThreeWayRowResult[] = req.rows.map((r) => {
    const fallbackRow =
      typeof r.rowNumber === 'number' && !Number.isNaN(r.rowNumber) ? Math.max(1, Math.floor(r.rowNumber)) : null;
    const baseRowNumber =
      typeof r.baseRowNumber === 'number' && !Number.isNaN(r.baseRowNumber) ? Math.max(1, Math.floor(r.baseRowNumber)) : fallbackRow;
    const oursRowNumber =
      typeof r.oursRowNumber === 'number' && !Number.isNaN(r.oursRowNumber) ? Math.max(1, Math.floor(r.oursRowNumber)) : fallbackRow;
    const theirsRowNumber =
      typeof r.theirsRowNumber === 'number' && !Number.isNaN(r.theirsRowNumber) ? Math.max(1, Math.floor(r.theirsRowNumber)) : fallbackRow;

    return {
      sheetName: resolvedSheetName,
      rowNumber: fallbackRow ?? undefined,
      baseRowNumber: baseRowNumber ?? null,
      oursRowNumber: oursRowNumber ?? null,
      theirsRowNumber: theirsRowNumber ?? null,
      colCount,
      base: readRowAligned(baseWs, baseRowNumber ?? null, 'base'),
      ours: readRowAligned(oursWs, oursRowNumber ?? null, 'ours'),
      theirs: readRowAligned(theirsWs, theirsRowNumber ?? null, 'theirs'),
    };
  });

  return { sheetName: resolvedSheetName, colCount, rows };
});
