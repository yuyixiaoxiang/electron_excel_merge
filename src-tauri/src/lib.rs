use serde::Serialize;
use std::collections::BTreeMap;
use std::path::PathBuf;

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct ProcessContext {
  args: Vec<String>,
  current_dir: Option<String>,
  env: BTreeMap<String, String>,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct CliThreeWayInfo {
  base_path: String,
  ours_path: String,
  theirs_path: String,
  merged_path: Option<String>,
  merged_path_raw: Option<String>,
  mode: String,
}

fn collect_process_context() -> ProcessContext {
  let current_dir = std::env::current_dir()
    .ok()
    .map(|path| path.to_string_lossy().into_owned());

  let mut env = BTreeMap::new();
  for key in ["PWD", "INIT_CWD", "OLDPWD", "GIT_WORK_TREE", "PATH", "Path"] {
    if let Ok(value) = std::env::var(key) {
      env.insert(key.to_string(), value);
    }
  }

  ProcessContext {
    args: std::env::args().collect(),
    current_dir,
    env,
  }
}

fn serialize_for_script<T: Serialize>(value: &T) -> String {
  serde_json::to_string(value)
    .unwrap_or_else(|_| "null".to_string())
    .replace('<', "\\u003c")
    .replace('>', "\\u003e")
    .replace('&', "\\u0026")
}

fn strip_outer_quotes(value: &str) -> String {
  let trimmed = value.trim();
  let quote_pairs = [('"', '"'), ('\'', '\''), ('“', '”'), ('‘', '’')];
  for (start, end) in quote_pairs {
    if trimmed.starts_with(start) && trimmed.ends_with(end) && trimmed.len() >= start.len_utf8() + end.len_utf8() {
      let start_index = start.len_utf8();
      let end_index = trimmed.len() - end.len_utf8();
      return trimmed[start_index..end_index].trim().to_string();
    }
  }
  trimmed.to_string()
}

fn normalize_cli_path_text(value: &str) -> String {
  let raw = strip_outer_quotes(value);
  let bytes = raw.as_bytes();
  if bytes.len() >= 4 && bytes[0] == b'/' && bytes[2] == b'/' && bytes[1].is_ascii_alphabetic() {
    let drive = char::from(bytes[1]).to_ascii_uppercase();
    let suffix = raw[3..].replace('/', "\\");
    return format!("{drive}:\\{suffix}");
  }
  raw
}

fn is_excel_like_path_text(value: &str) -> bool {
  let normalized = normalize_cli_path_text(value).to_ascii_lowercase();
  normalized.ends_with(".xlsx") || normalized.ends_with(".xlsm") || normalized.ends_with(".xls")
}

fn resolve_msys_tmp_path(value: &str) -> Option<String> {
  let normalized = normalize_cli_path_text(value);
  let lowercase = normalized.to_ascii_lowercase();
  if lowercase != "/tmp" && !lowercase.starts_with("/tmp/") {
    return None;
  }

  let mut temp_path = std::env::temp_dir();
  let relative = normalized
    .strip_prefix("/tmp")
    .unwrap_or("")
    .trim_start_matches('/');
  for part in relative.split('/') {
    if !part.is_empty() {
      temp_path.push(part);
    }
  }
  Some(temp_path.to_string_lossy().into_owned())
}

fn resolve_cli_path(current_dir: Option<&str>, value: &str) -> String {
  if let Some(msys_temp_path) = resolve_msys_tmp_path(value) {
    return msys_temp_path;
  }
  let normalized = normalize_cli_path_text(value);
  let path = PathBuf::from(&normalized);
  if path.is_absolute() {
    return path.to_string_lossy().into_owned();
  }
  if let Some(cwd) = current_dir {
    return PathBuf::from(cwd).join(path).to_string_lossy().into_owned();
  }
  normalized
}

fn parse_cli_three_way_info(context: &ProcessContext) -> Option<CliThreeWayInfo> {
  let raw_args: Vec<String> = context.args.iter().skip(1).cloned().collect();
  let user_args: Vec<String> = raw_args
    .iter()
    .map(|arg| strip_outer_quotes(arg))
    .filter(|arg| !arg.is_empty() && !arg.starts_with("--") && arg != "-Embedding")
    .collect();
  let excel_args: Vec<String> = user_args
    .iter()
    .filter(|arg| is_excel_like_path_text(arg))
    .cloned()
    .collect();
  let candidate_args = if excel_args.len() >= 2 { excel_args } else { user_args };

  if candidate_args.len() == 2 {
    let ours_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[0]);
    let theirs_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[1]);
    return Some(CliThreeWayInfo {
      base_path: ours_path.clone(),
      ours_path,
      theirs_path,
      merged_path: None,
      merged_path_raw: None,
      mode: "diff".to_string(),
    });
  }

  if candidate_args.len() < 3 {
    return None;
  }

  if candidate_args.len() == 3 {
    let ours_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[0]);
    let theirs_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[1]);
    let merged_path_raw = Some(normalize_cli_path_text(&candidate_args[2]));
    let merged_path = merged_path_raw
      .as_deref()
      .map(|value| resolve_cli_path(context.current_dir.as_deref(), value));

    return Some(CliThreeWayInfo {
      base_path: ours_path.clone(),
      ours_path,
      theirs_path,
      merged_path,
      merged_path_raw,
      mode: "simple-merge".to_string(),
    });
  }

  let base_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[0]);
  let ours_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[1]);
  let theirs_path = resolve_cli_path(context.current_dir.as_deref(), &candidate_args[2]);
  let merged_path_raw = candidate_args.get(3).map(|arg| normalize_cli_path_text(arg));
  let merged_path = merged_path_raw
    .as_deref()
    .map(|value| resolve_cli_path(context.current_dir.as_deref(), value));

  Some(CliThreeWayInfo {
    base_path,
    ours_path,
    theirs_path,
    merged_path,
    merged_path_raw,
    mode: "merge".to_string(),
  })
}

fn build_process_context_init_script(context: &ProcessContext) -> String {
  let json = serialize_for_script(context);

  format!(
    r#"
(() => {{
  const value = {json};
  if (!Object.prototype.hasOwnProperty.call(globalThis, "__EMERGE_PROCESS_CONTEXT__")) {{
    Object.defineProperty(globalThis, "__EMERGE_PROCESS_CONTEXT__", {{
      value,
      writable: false,
      configurable: true
    }});
  }}
}})();
"#
  )
}

fn build_process_context_page_script(context: &ProcessContext) -> String {
  let json = serialize_for_script(context);

  format!(
    r#"
(() => {{
  const value = {json};
  if (!Object.prototype.hasOwnProperty.call(globalThis, "__EMERGE_PROCESS_CONTEXT__")) {{
    Object.defineProperty(globalThis, "__EMERGE_PROCESS_CONTEXT__", {{
      value,
      writable: false,
      configurable: true
    }});
  }}
  window.dispatchEvent(new CustomEvent("emerge:process-context-ready", {{ detail: value }}));
}})();
"#
  )
}

fn build_cli_info_page_script(info: &Option<CliThreeWayInfo>) -> String {
  let json = serialize_for_script(info);

  format!(
    r#"
(() => {{
  const value = {json};
  if (!Object.prototype.hasOwnProperty.call(globalThis, "__EMERGE_CLI_INFO__")) {{
    Object.defineProperty(globalThis, "__EMERGE_CLI_INFO__", {{
      value,
      writable: false,
      configurable: true
    }});
  }}
  window.dispatchEvent(new CustomEvent("emerge:cli-info-ready", {{ detail: value }}));
}})();
"#
  )
}

fn write_startup_trace(context: &ProcessContext) {
  let trace_path = std::env::temp_dir().join("eMerge-startup-context.json");
  let payload = serde_json::to_string_pretty(context)
    .unwrap_or_else(|_| "{\"error\":\"failed to serialize process context\"}".to_string());
  let _ = std::fs::write(trace_path, payload);
}

#[tauri::command]
fn get_process_context(context: tauri::State<'_, ProcessContext>) -> ProcessContext {
  context.inner().clone()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
  let process_context = collect_process_context();
  let cli_info = parse_cli_three_way_info(&process_context);
  let process_context_init_script = build_process_context_init_script(&process_context);
  let process_context_page_script = build_process_context_page_script(&process_context);
  let cli_info_page_script = build_cli_info_page_script(&cli_info);
  write_startup_trace(&process_context);

  tauri::Builder::default()
    .manage(process_context)
    .append_invoke_initialization_script(process_context_init_script)
    .plugin(tauri_plugin_dialog::init())
    .plugin(tauri_plugin_fs::init())
    .plugin(tauri_plugin_opener::init())
    .plugin(tauri_plugin_shell::init())
    .setup(|app| {
      if cfg!(debug_assertions) {
        app.handle().plugin(
          tauri_plugin_log::Builder::default()
            .level(log::LevelFilter::Info)
            .build(),
        )?;
      }
      Ok(())
    })
    .on_page_load(move |webview, _payload| {
      let _ = webview.eval(&process_context_page_script);
      let _ = webview.eval(&cli_info_page_script);
    })
    .invoke_handler(tauri::generate_handler![get_process_context])
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}
