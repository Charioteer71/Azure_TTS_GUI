import shutil
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, simpledialog, filedialog
import azure.cognitiveservices.speech as speechsdk
import threading
import xml.sax.saxutils
import json
import os
import time
import tempfile
import pygame

CONFIG_FILE_NAME = "azure_tts_settings.json"

class TextToSpeechApp:
    def __init__(self, master):
        self.master = master
        master.title("Azure 文本转语音 (v4.8.5 - 启动提示)") # 版本号和标题更新
        master.geometry("650x880")

        # --- 初始化样式 ---
        self.style = ttk.Style()
        self.style.configure(
            "Hint.TLabel",
            foreground="steel blue", # 提示文字颜色
            font=('Arial', 9, 'bold')  # 提示文字字体
        )
        # --- 结束初始化样式 ---

        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.cache_dir_path = os.path.join(self.script_dir, "azure_tts_cache")
        self._initialize_cache_directory()

        try:
            pygame.init()
            pygame.mixer.init()
            self.pygame_initialized = True
        except Exception as e:
            self.pygame_initialized = False
            messagebox.showerror("Pygame 初始化失败", f"Pygame mixer 初始化失败: {e}\n播放功能将受限或不可用。", parent=master)
            print(f"Pygame init error: {e}")

        self.config_file_path = os.path.join(self.script_dir, CONFIG_FILE_NAME)

        # Playback State & Cache
        self.playback_state = "idle"
        self.synthesized_audio_filepath = None
        self.total_audio_duration_sec = 0
        self.progress_updater_id = None
        self.is_user_seeking = False
        self.last_synthesis_params = {} # For caching
        self.playback_marker_sec = 0
        self.playback_start_time_monotonic = None
        self._text_modified_flag = False # To detect text area changes for cache

        # App state variables
        self.all_voices_in_region = []
        self.loaded_voices_credentials = {"key": None, "region": None}
        self.current_language_voice_infos = {}
        self.voice_profiles_data = {}
        self._profile_being_loaded_settings = None
        self._pending_profile_to_apply_after_load = None

        # --- UI Setup ---
        # Azure 配置
        self.azure_config_frame = ttk.LabelFrame(master, text="Azure 配置")
        self.azure_config_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(self.azure_config_frame, text="订阅密钥:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.subscription_key_entry = ttk.Entry(self.azure_config_frame, width=45, show="*")
        self.subscription_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(self.azure_config_frame, text="服务区域:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.service_region_entry = ttk.Entry(self.azure_config_frame, width=45)
        self.service_region_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(self.azure_config_frame, text=" (例如: eastus)").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        
        self.azure_buttons_frame = ttk.Frame(self.azure_config_frame)
        self.azure_buttons_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky="ew") # 让框架水平填充

        self.save_credentials_button = ttk.Button(self.azure_buttons_frame, text="保存凭据", command=self.save_credentials)
        self.save_credentials_button.pack(side="left", padx=(0, 5)) # 左0右5间距

        self.load_voices_button = ttk.Button(self.azure_buttons_frame, text="加载/刷新语音列表", command=self.load_voices_from_azure)
        self.load_voices_button.pack(side="left", padx=(0, 5)) # 左0右5间距

        # --- 新增：醒目的提示标签 ---
        self.load_voices_hint_label = ttk.Label(
            self.azure_buttons_frame,
            text="<-- 启动后请先点此加载语音", 
            style="Hint.TLabel"
        )
        self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w') # 左对齐
        # --- 结束新增 ---

        # 语音配置
        self.voice_config_frame = ttk.LabelFrame(master, text="语音配置")
        self.voice_config_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(self.voice_config_frame, text="选择语言:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.language_var = tk.StringVar(); self.language_var.trace_add("write", self._on_voice_params_changed_for_cache)
        self.language_combo = ttk.Combobox(self.voice_config_frame, textvariable=self.language_var, values=["zh-CN", "en-US", "ja-JP", "ko-KR", "fr-FR", "de-DE", "es-ES"], state="disabled", exportselection=False, width=30)
        self.language_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.language_combo.bind("<<ComboboxSelected>>", self.on_language_selected)

        ttk.Label(self.voice_config_frame, text="选择语音:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.voice_var = tk.StringVar(); self.voice_var.trace_add("write", self._on_voice_params_changed_for_cache)
        self.voice_combo = ttk.Combobox(self.voice_config_frame, textvariable=self.voice_var, state="disabled", exportselection=False, width=30)
        self.voice_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.voice_combo.bind("<<ComboboxSelected>>", self.on_voice_selected)

        ttk.Label(self.voice_config_frame, text="角色风格:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.role_var = tk.StringVar(); self.role_var.trace_add("write", self._on_voice_params_changed_for_cache)
        self.role_combo = ttk.Combobox(self.voice_config_frame, textvariable=self.role_var, state="disabled", exportselection=False, width=30)
        self.role_combo.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.voice_config_frame, text="说话风格:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.style_var = tk.StringVar(); self.style_var.trace_add("write", self._on_voice_params_changed_for_cache)
        self.style_combo = ttk.Combobox(self.voice_config_frame, textvariable=self.style_var, state="disabled", exportselection=False, width=30)
        self.style_combo.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.voice_config_frame, text="语速:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.rate_var = tk.DoubleVar(value=1.0)
        self.rate_var.trace_add("write", self._on_rate_var_changed_for_cache_and_display)

        self.rate_slider_frame = ttk.Frame(self.voice_config_frame)
        self.rate_slider_frame.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        self.rate_slider = ttk.Scale(
            self.rate_slider_frame,
            variable=self.rate_var,
            from_=0.5,
            to=3.0,
            orient="horizontal",
            length=200 
        )
        self.rate_slider.pack(side="left", fill="x", expand=True)
        self.rate_slider.config(state="disabled")

        self.rate_display_var = tk.StringVar(value="1.00x")
        self.rate_display_label = ttk.Label(self.rate_slider_frame, textvariable=self.rate_display_var, width=7, anchor="e")
        self.rate_display_label.pack(side="left", padx=(5,0))

        self.profile_management_frame = ttk.LabelFrame(master, text="语音配置文件")
        self.profile_management_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(self.profile_management_frame, text="选择配置:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(self.profile_management_frame, textvariable=self.profile_var, state="readonly", exportselection=False, width=30)
        self.profile_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.profile_combo.bind("<<ComboboxSelected>>", self.on_profile_combobox_selected)
        self.save_profile_button = ttk.Button(self.profile_management_frame, text="保存当前为新配置", command=self.save_current_settings_as_profile)
        self.save_profile_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        
        self.text_input_frame = ttk.LabelFrame(master, text="输入文本")
        self.text_input_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.text_area = scrolledtext.ScrolledText(self.text_input_frame, wrap=tk.WORD, height=8, undo=True) 
        self.text_area.pack(padx=5, pady=5, fill="both", expand=True)
        self.text_area.insert(tk.END, "你好，世界！")
        self.text_area.bind("<<Modified>>", self._on_text_area_modified_flag) 

        self.playback_control_frame = ttk.LabelFrame(master, text="播放控制")
        self.playback_control_frame.pack(padx=10, pady=10, fill="x")
        self.play_pause_button = ttk.Button(self.playback_control_frame, text="▶️ 播放", command=self._on_play_pause_button_click, state=tk.DISABLED)
        self.play_pause_button.pack(side="left", padx=5, pady=5)
        self.stop_button = ttk.Button(self.playback_control_frame, text="⏹️ 停止", command=self._on_stop_button_click, state=tk.DISABLED)
        self.stop_button.pack(side="left", padx=5, pady=5)
        self.time_label_var = tk.StringVar(value="00:00 / 00:00")
        self.time_label = ttk.Label(self.playback_control_frame, textvariable=self.time_label_var, width=15, anchor="w")
        self.time_label.pack(side="right", padx=5, pady=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Scale(self.playback_control_frame, variable=self.progress_var, from_=0, to=100, orient="horizontal", state=tk.DISABLED, command=self._on_scale_drag_changed)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.progress_bar.bind("<ButtonPress-1>", self._on_scale_press)
        self.progress_bar.bind("<ButtonRelease-1>", self._on_scale_release)

        self.main_button_frame = ttk.Frame(master)
        self.main_button_frame.pack(padx=10, pady=10, fill="x", anchor="s")
        self.save_mp3_button = ttk.Button(self.main_button_frame, text="保存为 MP3", command=self.save_text_to_mp3_thread, state="disabled")
        self.save_mp3_button.pack(side="left", padx=5, pady=5)
        self.status_label = ttk.Label(self.main_button_frame, text="状态: 请先加载语音列表或配置文件")
        self.status_label.pack(side="left", padx=5, pady=5)

        self.azure_config_frame.columnconfigure(1, weight=1)
        self.voice_config_frame.columnconfigure(1, weight=1)
        self.voice_config_frame.grid_rowconfigure(4, pad=5) 
        self.profile_management_frame.columnconfigure(1, weight=1)

        self.load_app_config()
        master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _initialize_cache_directory(self):
        print(f"Debug: 正在初始化缓存目录: {self.cache_dir_path}")
        try:
            if os.path.exists(self.cache_dir_path):
                print(f"Debug: 尝试删除已存在的缓存目录: {self.cache_dir_path}")
                shutil.rmtree(self.cache_dir_path) 
                print(f"Debug: 成功删除已存在的缓存目录。")
        except Exception as e:
            print(f"警告: 未能移除已存在的缓存目录 '{self.cache_dir_path}'。错误: {e}")
            print(f"警告: 将继续尝试创建/使用目录。如果移除失败，旧文件可能残留。")
        try:
            os.makedirs(self.cache_dir_path, exist_ok=True) 
            print(f"Debug: 缓存目录已确保存在于: {self.cache_dir_path}")
        except Exception as e:
            messagebox.showerror(
                "关键错误",
                f"无法创建缓存目录: {self.cache_dir_path}\n错误: {e}\n"
                "程序可能无法正常工作，语音合成功能将受影响。",
                parent=self.master
            )
            print(f"严重错误: 未能创建缓存目录 {self.cache_dir_path}。语音合成功能将失败。")


    def _on_rate_var_changed_for_cache_and_display(self, *args):
        try:
            rate_val = self.rate_var.get()
            self.rate_display_var.set(f"{rate_val:.2f}x")
        except tk.TclError: 
            self.rate_display_var.set("1.00x")
        self._on_voice_params_changed_for_cache(*args)

    def _on_text_area_modified_flag(self, event=None):
        self.text_modified_flag = True

    def _on_voice_params_changed_for_cache(self, *args):
        pass

    def _get_current_synthesis_params(self):
        return {
            "text": self.text_area.get("1.0", tk.END).strip(),
            "lang": self.language_var.get(),
            "voice": self.voice_var.get(),
            "role": self.role_var.get(),
            "style": self.style_var.get(),
            "rate": self.rate_var.get(),
            "subscription_key": self.subscription_key_entry.get(),
            "service_region": self.service_region_entry.get(),
        }

    def _update_status(self, message):
        if hasattr(self, 'status_label') and self.status_label and self.status_label.winfo_exists():
            self.status_label.config(text=f"状态: {message}")
            self.master.update_idletasks()

    def _get_default_config(self):
        return {"azure_credentials": {"subscription_key": "", "service_region": ""}, "voice_profiles": {}}

    def _update_ui_for_playback_state(self):
        if not self.pygame_initialized:
            self.play_pause_button.config(text="▶️ 播放", state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED); self.progress_bar.config(state=tk.DISABLED)
            self.save_mp3_button.config(state=tk.DISABLED); return

        lang_ok=bool(self.language_var.get()); voice_ok=bool(self.voice_var.get())
        voices_loaded=isinstance(self.all_voices_in_region,list) and bool(self.all_voices_in_region)
        text_present = bool(self.text_area.get("1.0", tk.END).strip())
        can_synthesize_new = lang_ok and voice_ok and voices_loaded and text_present
        
        self.save_mp3_button.config(state=tk.NORMAL if can_synthesize_new else tk.DISABLED)

        if self.playback_state == "idle" or self.playback_state == "stopped_by_user":
            can_play_cached = bool(self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath) and self.total_audio_duration_sec > 0)
            self.play_pause_button.config(text="▶️ 播放", state=tk.NORMAL if (can_synthesize_new or can_play_cached) else tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            self.progress_bar.config(state=tk.DISABLED if not can_play_cached else tk.NORMAL) 
            
            current_t_display = 0
            total_t_display = 0

            if can_play_cached: 
                total_t_display = self.total_audio_duration_sec
                if self.playback_state == "stopped_by_user": 
                    current_t_display = self.progress_var.get() 
            
            if self.playback_state == "idle" and not can_play_cached: 
                self.progress_var.set(0)

            self.time_label_var.set(f"{self._format_time(current_t_display)} / {self._format_time(total_t_display)}")
            if total_t_display > 0: 
                self.progress_bar.config(to=total_t_display)
            else: 
                self.progress_bar.config(to=100)
                self.progress_var.set(0)

        elif self.playback_state == "synthesizing":
            self.play_pause_button.config(text="合成中...", state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED) 
            self.progress_bar.config(state=tk.DISABLED)
        elif self.playback_state == "playing":
            self.play_pause_button.config(text="⏸️ 暂停", state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar.config(state=tk.NORMAL)
        elif self.playback_state == "paused":
            self.play_pause_button.config(text="▶️ 继续", state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar.config(state=tk.NORMAL)

    def load_app_config(self):
        try:
            if os.path.exists(self.config_file_path):
                with open(self.config_file_path, 'r', encoding='utf-8') as f: config_data = json.load(f)
            else:
                config_data = self._get_default_config(); self._update_status("未找到配置文件，使用默认。")
            credentials = config_data.get("azure_credentials", self._get_default_config()["azure_credentials"])
            self.subscription_key_entry.delete(0, tk.END); self.subscription_key_entry.insert(0, credentials.get("subscription_key", ""))
            self.service_region_entry.delete(0, tk.END); self.service_region_entry.insert(0, credentials.get("service_region", ""))
            self.voice_profiles_data = config_data.get("voice_profiles", self._get_default_config()["voice_profiles"])
            self._update_profile_combobox()
        except Exception as e:
            messagebox.showerror("加载配置错误", f"加载配置文件时出错: {e}。\n将使用默认设置。", parent=self.master)
            self._apply_default_config_ui()
        finally:
            self._update_ui_for_playback_state()

    def _apply_default_config_ui(self):
        default_config = self._get_default_config()
        self.subscription_key_entry.delete(0, tk.END); self.subscription_key_entry.insert(0, default_config["azure_credentials"]["subscription_key"])
        self.service_region_entry.delete(0, tk.END); self.service_region_entry.insert(0, default_config["azure_credentials"]["service_region"])
        self.voice_profiles_data = default_config["voice_profiles"]; self._update_profile_combobox()
        self.rate_var.set(1.0) 
        self._update_status("已应用默认配置。")

    def save_app_config(self):
        current_key = self.subscription_key_entry.get(); current_region = self.service_region_entry.get()
        config_data = self._get_default_config() 
        try: 
            if os.path.exists(self.config_file_path):
                with open(self.config_file_path, 'r', encoding='utf-8') as f: 
                    loaded_data = json.load(f)
                    config_data.update(loaded_data) 
        except (json.JSONDecodeError, FileNotFoundError):
            pass 
        
        config_data["azure_credentials"] = {"subscription_key": current_key, "service_region": current_region}
        config_data["voice_profiles"] = self.voice_profiles_data 
        try:
            with open(self.config_file_path, 'w', encoding='utf-8') as f: json.dump(config_data, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror("保存配置错误", f"无法写入配置文件: {e}", parent=self.master); return False

    def _clear_all_voice_data_and_ui(self):
        self.all_voices_in_region = []
        self.loaded_voices_credentials = {"key": None, "region": None}
        self.language_var.set("")
        # self.language_combo.set('') # Combobox doesn't have .set directly, use var
        self.language_combo.config(state="disabled", values=["zh-CN", "en-US", "ja-JP", "ko-KR", "fr-FR", "de-DE", "es-ES"])
        self._reset_voice_selections()

        # --- 修改：重新显示并重置提示标签的文本 ---
        if hasattr(self, 'load_voices_hint_label'):
            self.load_voices_hint_label.config(text="<-- 启动后请先点此加载语音") 
            if not self.load_voices_hint_label.winfo_ismapped(): 
                self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')
        # --- 结束修改 ---

    def save_credentials(self):
        current_key_in_field = self.subscription_key_entry.get()
        current_region_in_field = self.service_region_entry.get()

        if self.save_app_config(): 
            self._update_status("凭据和配置已保存。") 
            credentials_differ_from_loaded_voices = \
                (current_key_in_field != self.loaded_voices_credentials.get("key") or \
                 current_region_in_field != self.loaded_voices_credentials.get("region"))

            if credentials_differ_from_loaded_voices and self.all_voices_in_region:
                self.master.update_idletasks() 
                if messagebox.askyesno("凭据已更改",
                                     "保存的 Azure 凭据与当前加载的语音列表所用凭据不同。\n"
                                     "建议刷新语音列表以匹配新凭据。\n"
                                     "是否立即清除当前语音列表和相关选择？\n"
                                     "（之后您需要点击“加载/刷新语音列表”按钮）", parent=self.master):
                    self._clear_all_voice_data_and_ui() # This will re-show the hint
                    self._update_status("凭据已保存。语音列表已清除，请使用新凭据加载语音。")
                else:
                    self._update_status("凭据已保存。当前语音列表可能基于旧凭据，建议刷新。")
                    # Optionally, update hint text if not clearing immediately
                    if hasattr(self, 'load_voices_hint_label'):
                        self.load_voices_hint_label.config(text="<-- 凭据已改, 建议刷新")
                        if not self.load_voices_hint_label.winfo_ismapped():
                           self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')

            elif not self.all_voices_in_region: 
                self._update_status("凭据已保存。请点击“加载/刷新语音列表”以开始。")
                if hasattr(self, 'load_voices_hint_label'): # Ensure hint is visible if no voices loaded
                    self.load_voices_hint_label.config(text="<-- 请先加载语音")
                    if not self.load_voices_hint_label.winfo_ismapped():
                        self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')
        else:
            self._update_status("凭据/配置保存失败。")

    def _update_profile_combobox(self):
        profile_names = sorted(list(self.voice_profiles_data.keys()))
        self.profile_combo['values'] = profile_names
        current_selection = self.profile_var.get()
        if not profile_names: self.profile_var.set("")
        elif current_selection not in profile_names : self.profile_var.set(profile_names[0] if profile_names else "")

    def save_current_settings_as_profile(self):
        lang = self.language_var.get(); voice = self.voice_var.get()
        if not lang or not voice: messagebox.showwarning("无法保存", "请先选择有效的语言和语音。", parent=self.master); return
        profile_name = simpledialog.askstring("保存语音配置", "输入配置名称:", parent=self.master)
        if not profile_name or not profile_name.strip(): self._update_status("保存配置已取消或名称无效。"); return
        profile_name = profile_name.strip()
        if profile_name in self.voice_profiles_data and not messagebox.askyesno("确认覆盖", f"名为 '{profile_name}' 的配置已存在，要覆盖吗？", parent=self.master):
            self._update_status("覆盖操作已取消."); return
        current_settings = {
            "language": lang, "voice": voice,
            "role": self.role_var.get(), "style": self.style_var.get(),
            "rate": self.rate_var.get()
        }
        self.voice_profiles_data[profile_name] = current_settings
        if self.save_app_config():
            self._update_profile_combobox(); self.profile_var.set(profile_name)
            messagebox.showinfo("配置已保存", f"语音配置 '{profile_name}' 已保存。", parent=self.master); self._update_status(f"配置 '{profile_name}' 已保存。")
        else: 
            if profile_name in self.voice_profiles_data and self.voice_profiles_data[profile_name] == current_settings: del self.voice_profiles_data[profile_name] 
            self._update_status("保存语音配置失败。")

    def on_profile_combobox_selected(self, event=None):
        selected_profile_name = self.profile_var.get()
        if not selected_profile_name or selected_profile_name not in self.voice_profiles_data:
            return

        if not (isinstance(self.all_voices_in_region, list) and self.all_voices_in_region):
            self.master.update_idletasks()
            if messagebox.askyesno("需要加载语音",
                                 "需要先从 Azure 加载区域语音列表才能应用此配置。\n"
                                 "这将使用当前凭据字段中的密钥和区域。\n是否现在加载？",
                                 parent=self.master):
                self._pending_profile_to_apply_after_load = selected_profile_name
                self.load_voices_from_azure()
            else:
                self._update_status("加载配置文件已取消（需先加载区域语音）。")
            return

        settings_to_load = self.voice_profiles_data[selected_profile_name]
        self._update_status(f"正在加载配置: {selected_profile_name}...")
        self._profile_being_loaded_settings = settings_to_load.copy()

        self._cleanup_temp_file()
        self.last_synthesis_params = {}
        self.total_audio_duration_sec = 0
        self.text_modified_flag = True

        self.language_var.set(settings_to_load["language"])
        self.master.after(0, self.on_language_selected, None)


    def _reset_voice_selections(self):
        self.voice_var.set(""); self.voice_combo.config(values=[], state="disabled")
        self.role_var.set("(无)"); self.role_combo.config(values=[], state="disabled")
        self.style_var.set("(默认)"); self.style_combo.config(values=[], state="disabled")
        self.rate_var.set(1.0) 
        self.rate_slider.config(state="disabled") 
        self._update_ui_for_playback_state() 

    def on_language_selected(self, event=None):
        current_lang_selection = self.language_var.get()
        self._reset_voice_selections() 
        self.language_var.set(current_lang_selection) 
        if not current_lang_selection:
            if self._profile_being_loaded_settings: self._profile_being_loaded_settings = None
            self._update_ui_for_playback_state(); return
        if not (isinstance(self.all_voices_in_region, list) and self.all_voices_in_region):
            self._update_status(f"请先成功加载 '{current_lang_selection}' 的语音列表。")
            if self._profile_being_loaded_settings:
                messagebox.showerror("配置加载错误", "无法加载配置：Azure 语音列表当前不可用。", parent=self.master)
                self._profile_being_loaded_settings = None 
                self.rate_var.set(1.0) 
                self.rate_slider.config(state="disabled")
            self._update_ui_for_playback_state(); return
        
        self.current_language_voice_infos.clear()
        voice_short_names = []
        for vi_idx, vi in enumerate(self.all_voices_in_region):
            try:
                if not (hasattr(vi, 'locale') and hasattr(vi, 'short_name') and isinstance(vi.locale, str) and isinstance(vi.short_name, str)): continue 
                if vi.locale == current_lang_selection:
                    voice_short_names.append(vi.short_name); self.current_language_voice_infos[vi.short_name] = vi
            except Exception as e_vi_proc: print(f"Debug: Error processing VoiceInfo: {e_vi_proc}"); continue
        voice_short_names.sort()

        if voice_short_names:
            self.voice_combo.config(values=voice_short_names, state="readonly")
            default_voice_to_set = voice_short_names[0]
            if self._profile_being_loaded_settings and self._profile_being_loaded_settings["language"] == current_lang_selection:
                profile_voice = self._profile_being_loaded_settings["voice"]
                if profile_voice in voice_short_names: default_voice_to_set = profile_voice
                else:
                    messagebox.showwarning("配置加载警告", f"配置语音 '{profile_voice}' 在 '{current_lang_selection}' 下未找到。", parent=self.master)
                    if self._profile_being_loaded_settings: self._profile_being_loaded_settings["_voice_application_failed"] = True
            self.voice_var.set(default_voice_to_set) 
            self.on_voice_selected(event=None) 
        else:
            self.voice_var.set("") 
            self._update_status(f"语言 '{current_lang_selection}' 没有找到可用语音。")
            if self._profile_being_loaded_settings:
                messagebox.showwarning("配置加载警告", f"无法为配置 '{self.profile_var.get()}' 在语言 '{current_lang_selection}' 下找到语音。", parent=self.master)
                self._profile_being_loaded_settings = None 
                self.rate_var.set(1.0) 
                self.rate_slider.config(state="disabled")
            self.on_voice_selected(event=None)


    def on_voice_selected(self, event=None):
        selected_voice_name = self.voice_var.get()
        
        self.role_combo.config(values=["(无)"], state="disabled"); self.role_var.set("(无)")
        self.style_combo.config(values=["(默认)"], state="disabled"); self.style_var.set("(默认)")

        was_loading_profile = bool(self._profile_being_loaded_settings)
        voices_loaded_ok = isinstance(self.all_voices_in_region, list) and bool(self.all_voices_in_region)
        voice_info_available = selected_voice_name and selected_voice_name in self.current_language_voice_infos
        default_rate_to_set = 1.0

        if not (voices_loaded_ok and voice_info_available):
            status_detail = ""
            if not voices_loaded_ok: status_detail = "语音列表未就绪。"
            elif not selected_voice_name: status_detail = "未选择有效语音。"
            elif not voice_info_available: status_detail = f"无法找到语音 '{selected_voice_name}' 的详细信息。"
            final_status = f"语音选择处理失败: {status_detail}"
            
            self.rate_slider.config(state="disabled") 
            if was_loading_profile: 
                final_status = f"加载配置 '{self.profile_var.get()}' 部分失败: {status_detail}"
                if self._profile_being_loaded_settings:
                     default_rate_to_set = self._profile_being_loaded_settings.get("rate", 1.0)
                self._profile_being_loaded_settings = None 
            self.rate_var.set(default_rate_to_set) 
            self._update_status(final_status); self._update_ui_for_playback_state(); return 

        voice_info = self.current_language_voice_infos[selected_voice_name]
        sdk_roles, sdk_styles = [], []
        raw_roles_list = getattr(voice_info, 'role_play_list', None)
        if isinstance(raw_roles_list, list): sdk_roles = sorted([str(r).strip() for r in raw_roles_list if str(r).strip()])
        elif isinstance(raw_roles_list, str) and raw_roles_list: sdk_roles = sorted([r.strip() for r in raw_roles_list.split(',') if r.strip()])
        raw_styles_list = getattr(voice_info, 'style_list', None)
        if isinstance(raw_styles_list, list): sdk_styles = sorted([str(s).strip() for s in raw_styles_list if str(s).strip()])
        elif isinstance(raw_styles_list, str) and raw_styles_list: sdk_styles = sorted([s.strip() for s in raw_styles_list.split(',') if s.strip()])
        
        roles_to_display = ["(无)"] + sdk_roles; self.role_combo.config(values=roles_to_display, state="readonly" if sdk_roles else "disabled")
        default_role_to_set = roles_to_display[0]
        styles_to_display = ["(默认)"] + sdk_styles; self.style_combo.config(values=styles_to_display, state="readonly" if sdk_styles else "disabled")
        default_style_to_set = styles_to_display[0]

        if was_loading_profile and self._profile_being_loaded_settings and \
           not self._profile_being_loaded_settings.get("_voice_application_failed", False) and \
           self._profile_being_loaded_settings["voice"] == selected_voice_name:
            
            profile_role = self._profile_being_loaded_settings["role"]
            if profile_role in roles_to_display: default_role_to_set = profile_role
            else: messagebox.showwarning("配置加载警告", f"配置角色 '{profile_role}' 对语音 '{selected_voice_name}' 无效。", parent=self.master)
            
            profile_style = self._profile_being_loaded_settings["style"]
            if profile_style in styles_to_display: default_style_to_set = profile_style
            else: messagebox.showwarning("配置加载警告", f"配置风格 '{profile_style}' 对语音 '{selected_voice_name}' 无效。", parent=self.master)
            
            default_rate_to_set = self._profile_being_loaded_settings.get("rate", 1.0)

        if was_loading_profile: 
            self._profile_being_loaded_settings = None 
        
        self.role_var.set(default_role_to_set)
        self.style_var.set(default_style_to_set) 
        self.rate_var.set(default_rate_to_set)

        self.rate_slider.config(state="normal" if voice_info_available else "disabled")
        
        status_msg = f"语音: {selected_voice_name}"
        if not sdk_roles and not sdk_styles: status_msg += " (无额外角色/风格)"
        if was_loading_profile and not self._profile_being_loaded_settings and self.profile_var.get(): 
            status_msg = f"配置 '{self.profile_var.get()}' 已应用. {status_msg}"
        self._update_status(status_msg)
        self._update_ui_for_playback_state()


    def load_voices_from_azure(self):
        subscription_key = self.subscription_key_entry.get()
        service_region = self.service_region_entry.get()
        if not subscription_key or not service_region:
            messagebox.showerror("配置错误", "请输入有效的 Azure 订阅密钥和区域。", parent=self.master)
            # 确保提示在出错时也可见
            if hasattr(self, 'load_voices_hint_label') and not self.load_voices_hint_label.winfo_ismapped():
                self.load_voices_hint_label.config(text="<-- 请先配置并加载语音") 
                self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')
            return

        self._update_status("正在从 Azure 加载语音列表...")
        self.load_voices_button.config(state="disabled")
        self._clear_all_voice_data_and_ui() # This will show the hint
        self.language_combo.config(state="disabled")
        
        try:
            speech_config = speechsdk.SpeechConfig(subscription=subscription_key, region=service_region)
            temp_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=None)
            result = temp_synthesizer.get_voices_async().get()

            if result.reason == speechsdk.ResultReason.VoicesListRetrieved and result.voices:
                self.all_voices_in_region = result.voices
                self.loaded_voices_credentials = {"key": subscription_key, "region": service_region}
                self.language_combo.config(state="readonly")
                self._update_status("语音列表加载成功。请选择语言。")
                messagebox.showinfo("成功", f"成功加载 {len(self.all_voices_in_region)} 个区域语音。", parent=self.master)
                
                # 成功加载后隐藏提示
                if hasattr(self, 'load_voices_hint_label') and self.load_voices_hint_label.winfo_ismapped():
                    self.load_voices_hint_label.pack_forget()

                current_lang_after_load = self.language_var.get()
                if self._pending_profile_to_apply_after_load:
                    p_name = self._pending_profile_to_apply_after_load
                    self._pending_profile_to_apply_after_load = None
                    if p_name in self.profile_combo['values']:
                        self.profile_var.set(p_name)
                        self.on_profile_combobox_selected(event=None)
                    else:
                        self._update_status(f"待加载配置 '{p_name}' 未找到。")
                elif current_lang_after_load:
                    self.language_var.set(current_lang_after_load)
                    self.on_language_selected(event=None)
            else:
                self.all_voices_in_region = []
                self.loaded_voices_credentials = {"key": None, "region": None}
                error_msg_detail = result.cancellation_details.error_details if result and result.cancellation_details and result.cancellation_details.error_details else ""
                error_msg = f"获取语音列表失败: {result.reason if result else '未知'}" + (f" (详情: {error_msg_detail})" if error_msg_detail else "")
                messagebox.showerror("加载失败", error_msg, parent=self.master)
                self._update_status(f"加载失败: {error_msg.splitlines()[0]}")
                # 加载失败时确保提示可见并更新文本
                if hasattr(self, 'load_voices_hint_label'):
                    self.load_voices_hint_label.config(text="<-- 加载失败, 请检查后重试")
                    if not self.load_voices_hint_label.winfo_ismapped():
                        self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')
        except Exception as e:
            self.all_voices_in_region = []
            self.loaded_voices_credentials = {"key": None, "region": None}
            messagebox.showerror("发生错误", f"加载语音列表时出错: {str(e)}", parent=self.master)
            self._update_status(f"加载错误: {str(e)}")
            # 发生异常时确保提示可见并更新文本
            if hasattr(self, 'load_voices_hint_label'):
                self.load_voices_hint_label.config(text="<-- 加载出错, 请重试")
                if not self.load_voices_hint_label.winfo_ismapped():
                    self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')
        finally:
            self.load_voices_button.config(state="normal")
            self._update_ui_for_playback_state()
            # 如果最终语音列表仍为空，也确保提示可见
            if not self.all_voices_in_region and hasattr(self, 'load_voices_hint_label'):
                self.load_voices_hint_label.config(text="<-- 启动后请先点此加载语音") 
                if not self.load_voices_hint_label.winfo_ismapped():
                    self.load_voices_hint_label.pack(side="left", padx=(0, 5), anchor='w')

    def _get_common_synthesis_inputs(self, for_playback=False):
        s_key=self.subscription_key_entry.get();s_reg=self.service_region_entry.get();
        txt=self.text_area.get("1.0",tk.END).strip()
        lang=self.language_var.get();voice=self.voice_var.get()
        ready_to_synthesize = all([s_key,s_reg,lang,voice,txt]) and isinstance(self.all_voices_in_region,list) and self.all_voices_in_region
        if not ready_to_synthesize:
            if not for_playback: 
                messagebox.showerror("输入错误","操作前，请确保所有必填项都已填写，且语音列表已成功加载。",parent=self.master)
            return None
        return s_key,s_reg,txt,lang,voice

    def _build_ssml(self, text_to_speak_raw, lang, voice_name, role, style, rate):
        txt_esc = xml.sax.saxutils.escape(text_to_speak_raw)
        parts = [
            f'<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xml:lang="{lang}">',
            f'<voice name="{voice_name}">'
        ]
        prosody_opened = False
        if abs(rate - 1.0) > 0.001: 
            rate_value_str = f"{rate:.2f}" 
            parts.append(f'<prosody rate="{rate_value_str}">')
            prosody_opened = True
        expr_as_opened = False
        if role != "(无)" or style != "(默认)":
            attrs = []
            if role != "(无)": attrs.append(f'role="{role}"')
            if style != "(默认)": attrs.append(f'style="{style}"')
            if attrs:
                parts.append(f'<mstts:express-as {" ".join(attrs)}>')
                expr_as_opened = True
        parts.append(txt_esc)
        if expr_as_opened: parts.append('</mstts:express-as>')
        if prosody_opened: parts.append('</prosody>')
        parts.extend(['</voice>', '</speak>'])
        return "".join(parts)

    def _on_closing(self):
        if self.pygame_initialized and pygame.mixer.get_init(): 
            try:
                pygame.mixer.music.stop()
                pygame.mixer.music.unload()
            except pygame.error as e:
                print(f"Debug: 在 _on_closing 中停止/卸载 Pygame 音频时出错 (可忽略): {e}")
            except Exception as e_pg_close:
                 print(f"Debug: 在 _on_closing 中 Pygame 关闭操作时发生意外错误: {e_pg_close}")
        self._cleanup_temp_file() 
        if self.pygame_initialized:
            try:
                pygame.mixer.quit() 
                pygame.quit()       
            except Exception as e:
                print(f"Debug: Pygame 退出时发生错误: {e}")
        self.master.destroy()

    def _cleanup_temp_file(self):
        if not self.synthesized_audio_filepath or not os.path.exists(self.synthesized_audio_filepath):
            if self.synthesized_audio_filepath and not os.path.exists(self.synthesized_audio_filepath):
                print(f"Debug: 记录的临时文件 {self.synthesized_audio_filepath} 已不存在, 清除路径。")
            self.synthesized_audio_filepath = None
            return

        filepath_to_delete = self.synthesized_audio_filepath
        if self.pygame_initialized and pygame.mixer.get_init():
            try:
                pygame.mixer.music.stop()
                pygame.mixer.music.unload()
                time.sleep(0.05) 
            except pygame.error as e:
                print(f"Debug: 在 _cleanup_temp_file 中为 {filepath_to_delete} 停止/卸载 Pygame 音频时出错 (可忽略): {e}")
            except Exception as e_pg:
                print(f"Debug: 在 _cleanup_temp_file 中为 {filepath_to_delete} 处理 Pygame 时发生意外错误: {e_pg}")
        
        deleted_successfully = False
        for attempt in range(3):
            try:
                os.remove(filepath_to_delete)
                print(f"Debug: 成功清理临时文件: {filepath_to_delete}")
                deleted_successfully = True
                break 
            except PermissionError as e_perm:
                print(f"Debug: 清理尝试 {attempt + 1} 删除 '{filepath_to_delete}' 失败 (PermissionError): {e_perm}")
                if attempt < 2: 
                    time.sleep(0.1 * (attempt + 1)) 
                else: 
                    print(f"错误: 多次尝试后未能删除临时文件 '{filepath_to_delete}'。它可能被其他进程或 Pygame 锁定。")
            except FileNotFoundError:
                print(f"Debug: 临时文件 '{filepath_to_delete}' 已被删除或未找到。")
                deleted_successfully = True 
                break
            except Exception as e_os_rem: 
                print(f"错误: 删除 '{filepath_to_delete}' 时发生意外的操作系统错误: {e_os_rem}")
                break 
        if deleted_successfully:
            if self.synthesized_audio_filepath == filepath_to_delete:
                self.synthesized_audio_filepath = None

    def _format_time(self, seconds):
        if seconds is None or seconds < 0: return "00:00"
        minutes = int(seconds // 60); seconds = int(seconds % 60)
        return f"{minutes:02d}:{seconds:02d}"
    
    def _synthesize_audio_to_file_thread(self):
        current_params = self._get_current_synthesis_params()
        common_inputs = self._get_common_synthesis_inputs(for_playback=True) 
        if not common_inputs:
            self._update_status("输入错误，无法开始合成。")
            self.playback_state = "idle"
            self.master.after(0, self._update_ui_for_playback_state)
            return

        self.playback_state = "synthesizing"
        self.master.after(0, self._update_ui_for_playback_state)
        self._update_status("正在合成语音...")
        
        s_key, s_reg, txt_raw, lang, voice = common_inputs 
        role, style_val = current_params["role"], current_params["style"]
        rate_val = current_params["rate"] 
        ssml = self._build_ssml(txt_raw, lang, voice, role, style_val, rate_val)
        self._cleanup_temp_file() 
        
        try:
            fd, temp_path = tempfile.mkstemp(suffix=".wav", prefix="azure_tts_", dir=self.cache_dir_path)
            os.close(fd) 
            self.synthesized_audio_filepath = temp_path 
            print(f"Debug: 创建新的临时音频文件于: {self.synthesized_audio_filepath}")
            
            speech_config_obj = speechsdk.SpeechConfig(subscription=s_key, region=s_reg) 
            speech_config_obj.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Riff16Khz16BitMonoPcm)
            audio_config_obj = speechsdk.audio.AudioOutputConfig(filename=self.synthesized_audio_filepath) 
            synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config_obj, audio_config=audio_config_obj)
            result = synthesizer.speak_ssml_async(ssml).get()
            del synthesizer 

            if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
                self.total_audio_duration_sec = result.audio_duration.total_seconds() if result.audio_duration else 0
                self.last_synthesis_params = current_params 
                self.text_modified_flag = False 
                self._update_status("合成完毕，准备播放。")
                self.master.after(0, self._start_playback_after_synthesis, True) 
            else: 
                details = result.cancellation_details if result else None
                error_message_detail = ""
                if details:
                    error_message_detail = f"错误原因: {details.reason}"
                    if details.reason == speechsdk.CancellationReason.Error and details.error_details:
                        error_message_detail += f" - 错误详情: {details.error_details}"
                msg = f"语音合成取消/失败: {result.reason if result else '未知'}\n{error_message_detail}"
                self.master.after(0, lambda m=msg: messagebox.showerror("合成错误", m, parent=self.master))
                self.playback_state = "idle"
                self._cleanup_temp_file() 
                self.master.after(0, self._update_ui_for_playback_state)
                self.master.after(0, lambda r=result.reason: self._update_status(f"合成错误: {r if r else '未知'}"))
        except Exception as e: 
            self.master.after(0, lambda err=str(e): messagebox.showerror("发生严重错误", f"语音合成或文件操作失败: {err}", parent=self.master))
            self.playback_state = "idle"
            self._cleanup_temp_file() 
            self.master.after(0, self._update_ui_for_playback_state)
            self.master.after(0, lambda err=str(e): self._update_status(f"合成严重错误: {err}"))

    def _start_playback_after_synthesis(self, is_newly_synthesized=False):
        if self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath) and self.pygame_initialized:
            try:
                pygame.mixer.music.load(self.synthesized_audio_filepath)
                pygame.mixer.music.play()
                self.playback_state = "playing"
                self.playback_marker_sec = 0 
                self.playback_start_time_monotonic = time.monotonic()
                if self.total_audio_duration_sec > 0: self.progress_bar.config(to=self.total_audio_duration_sec)
                else: self.progress_bar.config(to=100)
                self.progress_var.set(0)
                self._schedule_progress_update()
            except pygame.error as e:
                messagebox.showerror("播放错误", f"无法播放音频文件: {e}", parent=self.master)
                self._update_status(f"播放错误: {e}"); self.playback_state = "idle"; self._cleanup_temp_file()
        else: self.playback_state = "idle" 
        self._update_ui_for_playback_state()

    def _on_play_pause_button_click(self):
        if not self.pygame_initialized: messagebox.showwarning("播放错误", "Pygame未能正确初始化。", parent=self.master); return
        current_params = self._get_current_synthesis_params()
        if self.playback_state == "playing": 
            if pygame.mixer.music.get_busy(): pygame.mixer.music.pause()
            self.playback_state = "paused"
            if self.progress_updater_id: self.master.after_cancel(self.progress_updater_id); self.progress_updater_id = None
            if self.playback_start_time_monotonic is not None:
                elapsed = time.monotonic() - self.playback_start_time_monotonic
                self.playback_marker_sec += elapsed
                self.playback_start_time_monotonic = None 
            self._update_status("已暂停。")
        elif self.playback_state == "paused": 
            pygame.mixer.music.unpause()
            self.playback_state = "playing"
            self.playback_start_time_monotonic = time.monotonic() 
            self._schedule_progress_update()
            self._update_status("继续播放...")
        elif self.playback_state == "idle" or self.playback_state == "stopped_by_user": 
            needs_resynthesis = True 
            if self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath) and \
               not self.text_modified_flag:
                if all(current_params.get(k) == self.last_synthesis_params.get(k) for k in ["lang", "voice", "role", "style", "rate", "subscription_key", "service_region"]) and \
                   current_params["text"] == self.last_synthesis_params.get("text"):
                    needs_resynthesis = False
            if not needs_resynthesis:
                self._update_status("播放已缓存音频...")
                self._start_playback_after_synthesis(is_newly_synthesized=False)
            else: 
                if not self._get_common_synthesis_inputs(for_playback=True):
                    self._update_status("输入不完整，无法播放。")
                    self.playback_state = "idle" 
                    self._update_ui_for_playback_state()
                    return
                self._update_status("准备合成新音频...")
                self.total_audio_duration_sec = 0; self.progress_var.set(0) 
                self.time_label_var.set("00:00 / 00:00")
                self.last_synthesis_params = {} 
                threading.Thread(target=self._synthesize_audio_to_file_thread, daemon=True).start()
        self._update_ui_for_playback_state()

    def _on_stop_button_click(self):
        if not self.pygame_initialized: return
        if self.progress_updater_id: 
            self.master.after_cancel(self.progress_updater_id)
            self.progress_updater_id = None
        if pygame.mixer.get_init() and pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
            pygame.mixer.music.unload() 
        self.playback_state = "stopped_by_user" 
        self.playback_marker_sec = 0 
        self.playback_start_time_monotonic = None 
        self.progress_var.set(0) 
        self.time_label_var.set(f"00:00 / {self._format_time(self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else 0)}")
        self._update_status("已停止。")
        self._update_ui_for_playback_state()

    def _schedule_progress_update(self):
        if self.progress_updater_id: self.master.after_cancel(self.progress_updater_id); self.progress_updater_id = None
        if self.playback_state == "playing" and self.pygame_initialized and pygame.mixer.music.get_busy() and not self.is_user_seeking and self.playback_start_time_monotonic is not None:
            elapsed_time_sec = time.monotonic() - self.playback_start_time_monotonic
            current_display_time_sec = self.playback_marker_sec + elapsed_time_sec
            current_display_time_sec = max(0, min(current_display_time_sec, self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else float('inf')))
            self.progress_var.set(current_display_time_sec)
            self.time_label_var.set(f"{self._format_time(current_display_time_sec)} / {self._format_time(self.total_audio_duration_sec)}")
            if self.total_audio_duration_sec > 0 and current_display_time_sec >= self.total_audio_duration_sec - 0.15: 
                self.master.after(0, self._on_stop_button_click) 
            else: self.progress_updater_id = self.master.after(100, self._schedule_progress_update)
        elif self.playback_state == "playing" and self.pygame_initialized and not pygame.mixer.music.get_busy():
            self.master.after(0, self._on_stop_button_click)

    def _on_scale_press(self, event):
        if self.playback_state in ["playing", "paused"] and self.total_audio_duration_sec > 0 and self.pygame_initialized and self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath):
            self.is_user_seeking = True
            if self.playback_state == "playing" and pygame.mixer.music.get_busy(): 
                pygame.mixer.music.pause()
                if self.playback_start_time_monotonic is not None: 
                    self.playback_marker_sec += (time.monotonic() - self.playback_start_time_monotonic)
                self.playback_start_time_monotonic = None 

    def _on_scale_release(self, event):
        if self.is_user_seeking and self.pygame_initialized and self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath):
            self.is_user_seeking = False
            seek_to_sec = self.progress_var.get()
            seek_to_sec = max(0, min(seek_to_sec, self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else 0))
            try:
                should_be_playing_after_seek = (self.playback_state == "playing") or \
                                               (self.playback_state == "paused" and self.play_pause_button.cget("text") == "▶️ 继续") 
                pygame.mixer.music.stop() 
                pygame.mixer.music.load(self.synthesized_audio_filepath) 
                pygame.mixer.music.play() 
                pygame.mixer.music.set_pos(seek_to_sec) 
                self.playback_marker_sec = seek_to_sec 
                self.playback_start_time_monotonic = time.monotonic()
                self.progress_var.set(seek_to_sec) 
                self.time_label_var.set(f"{self._format_time(seek_to_sec)} / {self._format_time(self.total_audio_duration_sec)}")
                if not should_be_playing_after_seek: 
                    pygame.mixer.music.pause()
                    self.playback_state = "paused"
                else: 
                    self.playback_state = "playing"
                    if pygame.mixer.music.get_busy() and not self.is_user_seeking : 
                        pass 
                    else: 
                        pygame.mixer.music.unpause() 
                    self._schedule_progress_update()
            except Exception as e:
                print(f"Error seeking audio: {e}")
                messagebox.showerror("播放错误", f"音频定位失败: {e}", parent=self.master)
            finally: 
                self.is_user_seeking = False
                self._update_ui_for_playback_state()
        elif self.is_user_seeking: 
            self.is_user_seeking = False
            self._update_ui_for_playback_state() 

    def _on_scale_drag_changed(self, value_str):
        if self.is_user_seeking and self.pygame_initialized and self.total_audio_duration_sec > 0:
            current_seek_sec = float(value_str)
            current_seek_sec = max(0, min(current_seek_sec, self.total_audio_duration_sec))
            self.time_label_var.set(f"{self._format_time(current_seek_sec)} / {self._format_time(self.total_audio_duration_sec)}")

    def save_text_to_mp3_thread(self):
        inputs = self._get_common_synthesis_inputs()
        if not inputs: 
            self._update_status("输入不完整，无法保存MP3。")
            return 
        self.play_pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.save_mp3_button.config(state=tk.DISABLED) 
        self.progress_bar.config(state=tk.DISABLED)
        self._update_status("准备保存 MP3...")
        threading.Thread(target=self.save_text_to_mp3, daemon=True).start()

    def save_text_to_mp3(self):
        inputs=self._get_common_synthesis_inputs()
        if not inputs: 
            self.master.after(0, lambda: self._update_status("输入错误，MP3保存取消。"))
            self.master.after(0, self._update_ui_for_playback_state) 
            return
        
        s_key, s_reg, txt_raw, lang, voice = inputs
        role, style_val = self.role_var.get(), self.style_var.get()
        rate_val = self.rate_var.get() 
        
        actual_filepath = filedialog.asksaveasfilename(
            defaultextension=".mp3", 
            filetypes=[("MP3 audio file","*.mp3"),("All files","*.*")], 
            title="保存 MP3 文件", 
            initialdir=self.script_dir, 
            parent=self.master 
        )
        
        if not actual_filepath: 
            self.master.after(0, lambda: self._update_status("MP3保存已取消"))
            self.master.after(0, self._update_ui_for_playback_state)
            return

        self.master.after(0, lambda p=actual_filepath: self._update_status(f"正在保存到 {os.path.basename(p)}..."))
        ssml = self._build_ssml(txt_raw, lang, voice, role, style_val, rate_val) 
        try:
            speech_config = speechsdk.SpeechConfig(subscription=s_key, region=s_reg)
            speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio16Khz64KBitRateMonoMp3)
            audio_config = speechsdk.audio.AudioOutputConfig(filename=actual_filepath)
            file_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
            result = file_synthesizer.speak_ssml_async(ssml).get()
            
            if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
                self.master.after(0, lambda p=actual_filepath: [
                    self._update_status(f"成功保存到 {os.path.basename(p)}"),
                    messagebox.showinfo("保存成功", f"语音已成功保存到:\n{p}", parent=self.master)
                ])
            elif result.reason == speechsdk.ResultReason.Canceled:
                details=result.cancellation_details
                msg=f"MP3保存取消: {details.reason}\n" + (f"错误详情: {details.error_details}" if details.reason==speechsdk.CancellationReason.Error and details.error_details else "")
                self.master.after(0, lambda m=msg, r=details.reason: [
                    messagebox.showerror("保存错误", m, parent=self.master),
                    self._update_status(f"MP3保存错误: {r}")
                ])
            else: 
                self.master.after(0, lambda r=result.reason: self._update_status(f"MP3保存遇到问题: {r}"))
            del file_synthesizer
        except Exception as e:
            self.master.after(0, lambda err=str(e): [
                messagebox.showerror("发生严重错误", f"MP3保存失败: {err}", parent=self.master),
                self._update_status(f"MP3保存严重错误: {err}")
            ])
        finally: 
            self.master.after(0, self._update_ui_for_playback_state)

if __name__ == "__main__":
    root = tk.Tk()
    app = TextToSpeechApp(root)
    root.mainloop()