use eframe::egui;
use std::collections::HashMap;
use std::sync::{Arc, Mutex, OnceLock};
use windows::Win32::{
    Foundation::{BOOL, HWND, LPARAM, LRESULT, WPARAM},
    System::{
        Com::{CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED},
        LibraryLoader::GetModuleHandleW,
    },
    UI::{
        Accessibility::{
            CUIAutomation, IUIAutomation, IUIAutomationElement, IUIAutomationTreeWalker,
            UIA_CONTROLTYPE_ID,
        },
        Input::KeyboardAndMouse::{
            GetKeyState, VK_CAPITAL, VK_CONTROL, VK_ESCAPE, VK_F1, VK_F12, VK_MENU, VK_RETURN,
            VK_SHIFT, VK_TAB,
        },
        WindowsAndMessaging::{
            CallNextHookEx, EnumWindows, GetForegroundWindow, GetWindowRect, GetWindowTextW,
            IsWindowVisible, SetWindowsHookExW, UnhookWindowsHookEx, HHOOK, KBDLLHOOKSTRUCT,
            MSLLHOOKSTRUCT, WH_KEYBOARD_LL, WH_MOUSE_LL, WM_KEYDOWN, WM_LBUTTONDOWN,
            WM_RBUTTONDOWN,
        },
    },
};

static mut TARGET_HWND: Option<HWND> = None; // Store the selected window's handle
static mut UI_AUTOMATION: Option<IUIAutomation> = None;
static mut KEY_LOGGER: Option<Arc<Mutex<String>>> = None;
static mut LAST_ACTION_WAS_KEYPRESS: bool = false;
static mut KEYBOARD_HOOK: Option<HHOOK> = None;
static mut MOUSE_HOOK: Option<HHOOK> = None;
static CONTROL_TYPE_MAP: OnceLock<HashMap<i32, &'static str>> = OnceLock::new();

struct MyApp {
    text: Arc<Mutex<String>>, // Shared across threads
    is_running: bool,
    selected_window: Option<String>,
    selected_hwnd: Option<HWND>, // Store the HWND of the selected window
    window_list: Vec<(String, HWND)>, // Store the window titles and HWNDs
}

impl Default for MyApp {
    fn default() -> Self {
        unsafe {
            // Initialize COM library for the thread
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();

            // Create the IUIAutomation instance using CoCreateInstance
            let automation: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).unwrap();

            // Store the instance in the global variable
            UI_AUTOMATION = Some(automation);
        }

        Self {
            text: Arc::new(Mutex::new(String::new())),
            is_running: false,
            selected_window: None,
            selected_hwnd: None,
            window_list: get_window_list(),
        }
    }
}

// Low-level mouse hook callback
extern "system" fn mouse_hook(n_code: i32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    if n_code >= 0 {
        unsafe {
            let foreground_hwnd = GetForegroundWindow();
            if let Some(target_hwnd) = TARGET_HWND {
                if foreground_hwnd == target_hwnd {
                    let ms_struct = *(l_param.0 as *mut MSLLHOOKSTRUCT);
                    let x = ms_struct.pt.x;
                    let y = ms_struct.pt.y;

                    let mut log_action = String::new();

                    if w_param.0 == WM_LBUTTONDOWN as usize || w_param.0 == WM_RBUTTONDOWN as usize
                    {
                        if click_inside_window(x, y, target_hwnd) {
                            if let Some(element) = get_ui_element_at_point(x, y) {
                                // we need a cache of elements so that we can use get_cached_parents

                                let max_depth = 3;
                                let parents = get_parents(Some(element), max_depth);

                                // Generate an XPath-like structure for the element hierarchy
                                let xpath = generate_xpath(&parents);
                                println!("Generated XPath: {}", xpath);

                                if let Some(first_element) = parents.first() {
                                    match first_element.CurrentControlType() {
                                        Ok(control_type) => {
                                            println!(
                                                "Element control type: {:?}",
                                                get_control_type_name(control_type)
                                            );

                                            let action = match w_param.0 as u32 {
                                                WM_LBUTTONDOWN => "Left Click",
                                                WM_RBUTTONDOWN => "Right Click",
                                                _ => "Unknown Action",
                                            };
                                            println!("{} at ({}, {})", action, x, y);

                                            log_action
                                                .push_str(&format!("{}\t\t{}", action, xpath));
                                        }
                                        Err(e) => {
                                            log_action.push_str(&format!(
                                                "Failed to get control type: {:?}",
                                                e
                                            ));
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if !log_action.is_empty() {
                        if let Some(ref key_logger) = KEY_LOGGER {
                            let mut log = key_logger.lock().unwrap();

                            if LAST_ACTION_WAS_KEYPRESS {
                                log.push('\''); // Add ' end of string to complete the syntax started in keyboard hook
                            }

                            if !log.is_empty() {
                                log.push('\n');
                            }

                            log.push_str(&log_action); // Log the mouse click on element

                            LAST_ACTION_WAS_KEYPRESS = false;
                        }
                    }
                }
            }
        }
    }

    unsafe { CallNextHookEx(None, n_code, w_param, l_param) }
}

// Low-level keyboard hook callback
extern "system" fn keyboard_hook(n_code: i32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    if n_code >= 0 && w_param.0 == WM_KEYDOWN as usize {
        unsafe {
            let foreground_hwnd = GetForegroundWindow();
            if let Some(target_hwnd) = TARGET_HWND {
                if foreground_hwnd == target_hwnd {
                    let kb_struct = *(l_param.0 as *mut KBDLLHOOKSTRUCT);
                    let vk_code = kb_struct.vkCode as i32;

                    let shift_pressed = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                    let caps_lock_on = (GetKeyState(VK_CAPITAL.0 as i32) & 1) != 0;

                    let key = translate_keycode(vk_code, shift_pressed, caps_lock_on);

                    if let Some(ref key_logger) = KEY_LOGGER {
                        let mut log = key_logger.lock().unwrap();

                        // TODO fix this, not correct AI POOPOO
                        if key.len() == 1 && key.chars().next().unwrap().is_alphanumeric() {
                            log.push_str(&key);
                        } else {
                            log.push_str(&format!("\nPress Key '{}'", key));
                        }

                        println!("Pressed key: {}", key);
                    }

                    LAST_ACTION_WAS_KEYPRESS = true;
                }
            }
        }
    }

    unsafe { CallNextHookEx(None, n_code, w_param, l_param) }
}

// Callback to enumerate windows and store their titles and HWNDs
extern "system" fn enum_window_callback(hwnd: HWND, lparam: LPARAM) -> BOOL {
    unsafe {
        // Buffer for the window title
        let mut buffer: [u16; 512] = [0; 512];

        // Only add visible windows
        if IsWindowVisible(hwnd).as_bool() && GetWindowTextW(hwnd, &mut buffer) > 0 {
            let window_title = String::from_utf16_lossy(&buffer);
            let windows_list = &mut *(lparam.0 as *mut Vec<(String, HWND)>);
            windows_list.push((window_title.trim_matches(char::from(0)).to_string(), hwnd));
        }
    }
    BOOL(1) // Continue enumeration
}

fn click_inside_window(x: i32, y: i32, hwnd: HWND) -> bool {
    unsafe {
        if hwnd.0 == std::ptr::null_mut() {
            // No foreground window
            return false;
        }

        // Get the window rectangle
        let mut rect = std::mem::zeroed();
        if GetWindowRect(hwnd, &mut rect).is_err() {
            // Failed to get the rectangle
            return false;
        }

        // TODO: Remove this logging? Or put the bool into a variable for cleaner code
        println!(
            "Click was was inside window? {}",
            x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom
        );

        x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom
    }
}

//  Get all window titles and their HWNDs
fn get_window_list() -> Vec<(String, HWND)> {
    let mut window_list: Vec<(String, HWND)> = Vec::new();
    unsafe {
        _ = EnumWindows(
            Some(enum_window_callback),
            LPARAM(&mut window_list as *mut _ as isize),
        );
    }
    window_list
}

fn initialize_control_type_map() -> HashMap<i32, &'static str> {
    HashMap::from([
        (50000, "Button"),
        (50001, "Calendar"),
        (50002, "CheckBox"),
        (50003, "ComboBox"),
        (50004, "Edit"),
        (50005, "Hyperlink"),
        (50006, "Image"),
        (50007, "ListItem"),
        (50008, "List"),
        (50009, "Menu"),
        (50010, "MenuBar"),
        (50011, "MenuItem"),
        (50012, "ProgressBar"),
        (50013, "RadioButton"),
        (50014, "ScrollBar"),
        (50015, "Slider"),
        (50016, "Spinner"),
        (50017, "StatusBar"),
        (50018, "Tab"),
        (50019, "TabItem"),
        (50020, "Text"),
        (50021, "ToolBar"),
        (50022, "ToolTip"),
        (50023, "Tree"),
        (50024, "TreeItem"),
        (50025, "Custom"),
        (50026, "Group"),
        (50027, "Thumb"),
        (50028, "DataGrid"),
        (50029, "DataItem"),
        (50030, "Document"),
        (50031, "SplitButton"),
        (50032, "Window"),
        (50033, "Pane"),
        (50034, "Header"),
        (50035, "HeaderItem"),
        (50036, "Table"),
        (50037, "TitleBar"),
        (50038, "Separator"),
    ])
}

fn get_control_type_name(control_type: UIA_CONTROLTYPE_ID) -> &'static str {
    // Ensure the map is initialized and retrieve it
    let map = CONTROL_TYPE_MAP.get_or_init(initialize_control_type_map);

    // Look up the control type ID in the map, returning "Unknown" if not found
    map.get(&control_type.0).copied().unwrap_or("Unknown")
}

unsafe fn get_ui_element_at_point(x: i32, y: i32) -> Option<IUIAutomationElement> {
    if let Some(ref ui_automation) = UI_AUTOMATION {
        let result = ui_automation.ElementFromPoint(windows::Win32::Foundation::POINT { x, y });

        // Check if result is Ok and return the element
        if let Ok(element) = result {
            return Some(element);
        }
    }
    None
}

fn translate_keycode(vk_code: i32, shift_pressed: bool, caps_lock_on: bool) -> String {
    match vk_code {
        code if code == VK_SHIFT.0 as i32 => "[SHIFT]".to_string(),
        code if code == VK_RETURN.0 as i32 => "[ENTER]".to_string(),
        code if code == VK_TAB.0 as i32 => "[TAB]".to_string(),
        code if code == VK_CONTROL.0 as i32 => "[CTRL]".to_string(),
        code if code == VK_MENU.0 as i32 => "[ALT]".to_string(),
        code if code == VK_CAPITAL.0 as i32 => "[CAPS_LOCK]".to_string(),
        code if code == VK_ESCAPE.0 as i32 => "[ESC]".to_string(),
        code if (VK_F1.0 as i32..=VK_F12.0 as i32).contains(&code) => {
            format!("[F{}]", code - VK_F1.0 as i32 + 1)
        }
        _ => {
            let char = vk_code as u8 as char;
            if char.is_ascii_alphanumeric() {
                if shift_pressed ^ caps_lock_on {
                    char.to_ascii_uppercase().to_string()
                } else {
                    char.to_ascii_lowercase().to_string()
                }
            } else {
                format!("[UNKNOWN: {}]", vk_code)
            }
        }
    }
}

fn create_tree_walker(
    automation: &IUIAutomation,
) -> windows::core::Result<IUIAutomationTreeWalker> {
    // Create a condition for the TreeWalker.
    // This example uses a condition to match all elements.
    let true_condition = unsafe { automation.CreateTrueCondition()? };

    // Create a TreeWalker with the condition.
    let tree_walker = unsafe { automation.CreateTreeWalker(&true_condition)? };
    Ok(tree_walker)
}

unsafe fn get_parents(
    element: Option<IUIAutomationElement>,
    max_depth: usize,
) -> Vec<IUIAutomationElement> {
    let mut parents = Vec::new();
    let mut current_element = element;
    let mut current_depth = 0;

    if let Some(ref ui_automation) = UI_AUTOMATION {
        let tree_walker: Option<IUIAutomationTreeWalker> = create_tree_walker(ui_automation).ok();

        while let Some(element) = &current_element {
            println!("Iteration of parent search: {}", current_depth);
            parents.push(element.clone());
            current_depth += 1;

            if current_depth >= max_depth {
                break;
            }

            if let Some(ref walker) = tree_walker {
                current_element = walker.GetParentElement(element).ok();
            }
        }
    } else {
        println!("No UI_AUTOMATION, this is bad...")
    }

    parents
}

unsafe fn generate_xpath(parents: &[IUIAutomationElement]) -> String {
    parents
        .iter()
        .rev()
        .enumerate()
        .map(|(index, element)| {
            let automation_id = element.CurrentAutomationId().unwrap_or_default();
            let name = element.CurrentName().unwrap_or_default();
            let control_type = get_control_type_name(
                element
                    .CurrentControlType()
                    .unwrap_or(UIA_CONTROLTYPE_ID(0)),
            );

            if !automation_id.is_empty() {
                format!("/{}[@AutomationId=\"{}\"]", control_type, automation_id)
            } else if !name.is_empty() {
                format!("/{}[@Name=\"{}\"]", control_type, name)
            } else {
                format!("/{}[Index=\"{}\"]", control_type, index) // No Name or Automation id, maybe we should not push this at all and break future from writing? 
            }
            // Other options that could be used here are ClassName and HelpText
            // According to https://gdatasoftwareag.github.io/robotframework-flaui/keywords/3.6.1.html
        })
        .collect::<String>()
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Test creator");

            ui.horizontal(|ui| {
                if ui
                    .button(if self.is_running { "Stop" } else { "Start" })
                    .clicked()
                {
                    if self.is_running {
                        // Stop logic
                        unsafe {
                            if let Some(hook) = KEYBOARD_HOOK {
                                _ = UnhookWindowsHookEx(hook);
                                KEYBOARD_HOOK = None;
                            }
                            if let Some(hook) = MOUSE_HOOK {
                                _ = UnhookWindowsHookEx(hook);
                                MOUSE_HOOK = None;
                            }
                        }
                        self.is_running = false;
                    } else {
                        // Start logic
                        if let Some(hwnd) = self.selected_hwnd {
                            unsafe {
                                let text_arc = Arc::clone(&self.text);
                                KEY_LOGGER = Some(text_arc);

                                TARGET_HWND = Some(hwnd);

                                let h_instance = GetModuleHandleW(None).unwrap();

                                KEYBOARD_HOOK = SetWindowsHookExW(
                                    WH_KEYBOARD_LL,
                                    Some(keyboard_hook),
                                    h_instance,
                                    0,
                                )
                                .ok();

                                MOUSE_HOOK =
                                    SetWindowsHookExW(WH_MOUSE_LL, Some(mouse_hook), h_instance, 0)
                                        .ok();
                            }
                            self.is_running = true;
                        }
                    }
                }

                ui.label(format!(
                    "Currently: {}",
                    if self.is_running {
                        "Running"
                    } else {
                        "Stopped"
                    }
                ))
            });

            ui.horizontal(|ui| {
                // Dropdown to select the window to follow
                egui::ComboBox::from_label("Select Window")
                    .selected_text(
                        self.selected_window
                            .clone()
                            .unwrap_or_else(|| "None".to_string()),
                    )
                    .show_ui(ui, |ui| {
                        for (title, hwnd) in &self.window_list {
                            if ui
                                .selectable_value(
                                    &mut self.selected_window,
                                    Some(title.clone()),
                                    title,
                                )
                                .clicked()
                            {
                                self.selected_hwnd = Some(*hwnd);
                            }
                        }
                    });

                if ui.button("Refresh Windows").clicked() {
                    self.window_list = get_window_list();
                }
            });

            let mut text_guard = self.text.lock().unwrap();
            egui::ScrollArea::vertical()
                .auto_shrink([true, true])
                .show(ui, |ui| {
                    ui.add_sized([400.0, 200.0], egui::TextEdit::multiline(&mut *text_guard));
                });

            ui.horizontal(|ui| {
                if ui.button("Clear").clicked() {
                    text_guard.clear();
                    // TODO
                    // BUG: The first input is not registered after calling clear?
                    // Or is it because window was not in focus for the click, since the test-creator window was in focus to press clear
                    // Should maybe have confrimation first to avoid mistakes?
                }
                if ui.button("Save").clicked() {
                    // TODO
                    // Prompting a file saving option needs an external library like rfd
                    // However doing regular I/O file saving does not require a library as it's available through the STD
                    // Save the current self.text to file, maybe have a default location where a user could save somewhere?
                    // Like config
                }
            })
        });
    }
}

fn main() -> std::result::Result<(), eframe::Error> {
    let native_options = eframe::NativeOptions::default();
    eframe::run_native(
        "Test creator",
        native_options,
        Box::new(|_cc| Ok(Box::new(MyApp::default()))),
    )
}

/*
TODO

1. Left click functionality
2. Better keyboard tracking with Enter shift etc.
3. Utility buttons? Like adding a sleep / wait to the robot framework synxtax? Or change last line to an evaluator?

*/
