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
        master.title("Azure 文本转语音 (v4.8.3 - 缓存与进度条修复)")
        master.geometry("650x850")

        try:
            pygame.init()
            pygame.mixer.init()
            self.pygame_initialized = True
        except Exception as e:
            self.pygame_initialized = False
            messagebox.showerror("Pygame 初始化失败", f"Pygame mixer 初始化失败: {e}\n播放功能将受限或不可用。", parent=master)
            print(f"Pygame init error: {e}")

        self.script_dir = os.path.dirname(os.path.abspath(__file__))
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
        self.azure_buttons_frame.grid(row=2, column=0, columnspan=3, pady=5)
        self.save_credentials_button = ttk.Button(self.azure_buttons_frame, text="保存凭据", command=self.save_credentials) # This command will now be defined
        self.save_credentials_button.pack(side="left", padx=5)
        self.load_voices_button = ttk.Button(self.azure_buttons_frame, text="加载/刷新语音列表", command=self.load_voices_from_azure)
        self.load_voices_button.pack(side="left", padx=5)

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

        # 语音配置文件管理
        self.profile_management_frame = ttk.LabelFrame(master, text="语音配置文件")
        self.profile_management_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(self.profile_management_frame, text="选择配置:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(self.profile_management_frame, textvariable=self.profile_var, state="readonly", exportselection=False, width=30)
        self.profile_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.profile_combo.bind("<<ComboboxSelected>>", self.on_profile_combobox_selected)
        self.save_profile_button = ttk.Button(self.profile_management_frame, text="保存当前为新配置", command=self.save_current_settings_as_profile)
        self.save_profile_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        
        # 文本输入
        self.text_input_frame = ttk.LabelFrame(master, text="输入文本")
        self.text_input_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.text_area = scrolledtext.ScrolledText(self.text_input_frame, wrap=tk.WORD, height=8, undo=True) 
        self.text_area.pack(padx=5, pady=5, fill="both", expand=True)
        self.text_area.insert(tk.END, "你好，世界！")
        self.text_area.bind("<<Modified>>", self._on_text_area_modified_flag) 

        # 播放控制
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

        # 主操作按钮和状态
        self.main_button_frame = ttk.Frame(master)
        self.main_button_frame.pack(padx=10, pady=10, fill="x", anchor="s")
        self.save_mp3_button = ttk.Button(self.main_button_frame, text="保存为 MP3", command=self.save_text_to_mp3_thread, state="disabled")
        self.save_mp3_button.pack(side="left", padx=5, pady=5)
        self.status_label = ttk.Label(self.main_button_frame, text="状态: 请先加载语音列表或配置文件")
        self.status_label.pack(side="left", padx=5, pady=5)

        self.azure_config_frame.columnconfigure(1, weight=1)
        self.voice_config_frame.columnconfigure(1, weight=1)
        self.profile_management_frame.columnconfigure(1, weight=1)

        self.load_app_config()
        master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_text_area_modified_flag(self, event=None):
        self.text_modified_flag = True

    def _on_voice_params_changed_for_cache(self, *args):
        # If voice parameters (lang, voice, role, style) change,
        # the current synthesized audio is no longer valid for "play cached"
        # We don't need to clear the file immediately, but the next play will compare params.
        # Setting a general "params_changed" flag could be an option,
        # or just rely on the comparison in _on_play_pause_button_click.
        # For simplicity, we will compare all params directly in play click.
        # If we wanted to be more proactive:
        # if self.playback_state == "idle" or self.playback_state == "stopped_by_user":
        #     if self.synthesized_audio_filepath:
        #         self.last_synthesis_params = {} # Force re-synthesis on next play
        #         # Optionally, could also delete self.synthesized_audio_filepath here
        #         # and reset self.total_audio_duration_sec
        pass # Rely on direct comparison in _on_play_pause_button_click

    def _get_current_synthesis_params(self):
        return {
            "text": self.text_area.get("1.0", tk.END).strip(),
            "lang": self.language_var.get(),
            "voice": self.voice_var.get(),
            "role": self.role_var.get(),
            "style": self.style_var.get(),
            "subscription_key": self.subscription_key_entry.get(), # Needed for re-synthesis
            "service_region": self.service_region_entry.get(),   # Needed for re-synthesis
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
        # 文本必须非空才能进行新的合成操作
        text_present = bool(self.text_area.get("1.0", tk.END).strip())
        can_synthesize_new = lang_ok and voice_ok and voices_loaded and text_present
        
        # 保存MP3按钮的状态取决于是否可以进行新的合成
        self.save_mp3_button.config(state=tk.NORMAL if can_synthesize_new else tk.DISABLED)

        # 播放/暂停/停止按钮的逻辑
        if self.playback_state == "idle" or self.playback_state == "stopped_by_user":
            # 如果可以合成新音频，或者存在有效的已缓存音频，则“播放”按钮可用
            can_play_cached = bool(self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath) and self.total_audio_duration_sec > 0)
            self.play_pause_button.config(text="▶️ 播放", state=tk.NORMAL if (can_synthesize_new or can_play_cached) else tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            self.progress_bar.config(state=tk.DISABLED if not can_play_cached else tk.NORMAL) # 如果有缓存音频，进度条可操作（显示时长）
            
            current_t_display = 0
            total_t_display = 0

            if can_play_cached: # 如果有缓存文件，显示其时长
                total_t_display = self.total_audio_duration_sec
                if self.playback_state == "stopped_by_user": # 如果是用户停止的，可以保留上次的进度条位置
                    current_t_display = self.progress_var.get() 
                # else (idle) current_t_display 保持 0
            
            if self.playback_state == "idle" and not can_play_cached: # 如果是空闲且无缓存，进度条归零
                 self.progress_var.set(0)
                 # self.last_synthesis_params = {} # 在 v4.8.2 中，这个在 _on_play_pause_button_click 中处理

            self.time_label_var.set(f"{self._format_time(current_t_display)} / {self._format_time(total_t_display)}")
            if total_t_display > 0: 
                self.progress_bar.config(to=total_t_display)
            else: # 没有有效时长，进度条使用默认范围
                self.progress_bar.config(to=100)
                self.progress_var.set(0)


        elif self.playback_state == "synthesizing":
            self.play_pause_button.config(text="合成中...", state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED) # 合成时通常不允许停止（除非实现取消合成的逻辑）
            self.progress_bar.config(state=tk.DISABLED)
        elif self.playback_state == "playing":
            self.play_pause_button.config(text="⏸️ 暂停", state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar.config(state=tk.NORMAL)
        elif self.playback_state == "paused":
            self.play_pause_button.config(text="▶️ 继续", state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar.config(state=tk.NORMAL) # 暂停时也允许拖动进度条

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
        self._update_status("已应用默认配置。")

    def save_app_config(self):
        current_key = self.subscription_key_entry.get(); current_region = self.service_region_entry.get()
        config_data = self._get_default_config() # Start with a clean default structure
        try: # Try to load existing to preserve any other potential future keys, but overwrite known sections
            if os.path.exists(self.config_file_path):
                with open(self.config_file_path, 'r', encoding='utf-8') as f: 
                    loaded_data = json.load(f)
                    config_data.update(loaded_data) # Merge, new values will overwrite
        except (json.JSONDecodeError, FileNotFoundError):
            pass # Use default if file is bad or not found
        
        config_data["azure_credentials"] = {"subscription_key": current_key, "service_region": current_region}
        config_data["voice_profiles"] = self.voice_profiles_data # This is already up-to-date in memory
        try:
            with open(self.config_file_path, 'w', encoding='utf-8') as f: json.dump(config_data, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror("保存配置错误", f"无法写入配置文件: {e}", parent=self.master); return False

    def _clear_all_voice_data_and_ui(self):
        self.all_voices_in_region = []
        self.loaded_voices_credentials = {"key": None, "region": None}
        self.language_var.set(""); self.language_combo.set(''); 
        self.language_combo.config(state="disabled", values=["zh-CN", "en-US", "ja-JP", "ko-KR", "fr-FR", "de-DE", "es-ES"])
        self._reset_voice_selections()

    # >>> 定义 save_credentials 方法 <<<
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
                    self._clear_all_voice_data_and_ui() 
                    self._update_status("凭据已保存。语音列表已清除，请使用新凭据加载语音。")
                else:
                    self._update_status("凭据已保存。当前语音列表可能基于旧凭据，建议刷新。")
            elif not self.all_voices_in_region: 
                 self._update_status("凭据已保存。请点击“加载/刷新语音列表”以开始。")
        else:
            self._update_status("凭据/配置保存失败。")
    # >>> save_credentials 方法结束 <<<

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
        current_settings = {"language": lang, "voice": voice, "role": self.role_var.get(), "style": self.style_var.get()}
        self.voice_profiles_data[profile_name] = current_settings
        if self.save_app_config():
             self._update_profile_combobox(); self.profile_var.set(profile_name)
             messagebox.showinfo("配置已保存", f"语音配置 '{profile_name}' 已保存。", parent=self.master); self._update_status(f"配置 '{profile_name}' 已保存。")
        else: 
            if profile_name in self.voice_profiles_data and self.voice_profiles_data[profile_name] == current_settings: del self.voice_profiles_data[profile_name] # Rollback in-memory
            self._update_status("保存语音配置失败。")

    def on_profile_combobox_selected(self, event=None):
        selected_profile_name = self.profile_var.get()
        if not selected_profile_name or selected_profile_name not in self.voice_profiles_data: return
        if not (isinstance(self.all_voices_in_region, list) and self.all_voices_in_region):
            self.master.update_idletasks()
            if messagebox.askyesno("需要加载语音", "需要先从 Azure 加载区域语音列表才能应用此配置。\n这将使用当前凭据字段中的密钥和区域。\n是否现在加载？", parent=self.master):
                self._pending_profile_to_apply_after_load = selected_profile_name; self.load_voices_from_azure() 
            else: self._update_status("加载配置文件已取消（需先加载区域语音）。")
            return
        settings_to_load = self.voice_profiles_data[selected_profile_name]
        self._update_status(f"正在加载配置: {selected_profile_name}...")
        self._profile_being_loaded_settings = settings_to_load.copy()
        
        # Invalidate cache before applying profile, as params will change
        self._cleanup_temp_file()
        self.last_synthesis_params = {}
        self.total_audio_duration_sec = 0
        self.text_modified_flag = True # Force re-synthesis for the new profile

        if self.language_var.get() != settings_to_load["language"]: self.language_var.set(settings_to_load["language"]) # Triggers trace -> _on_voice_params_changed_for_cache
        self.on_language_selected(event=None) # Explicit call for cascade

    def _reset_voice_selections(self):
        self.voice_var.set(""); self.voice_combo.config(values=[], state="disabled")
        self.role_var.set("(无)"); self.role_combo.config(values=[], state="disabled")
        self.style_var.set("(默认)"); self.style_combo.config(values=[], state="disabled")
        self._update_ui_for_playback_state() 
        # self.save_mp3_button is handled by _update_ui_for_playback_state based on can_synthesize_new

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
                messagebox.showerror("配置加载错误", "无法加载配置：Azure 语音列表当前不可用。", parent=self.master); self._profile_being_loaded_settings = None
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
            self.voice_var.set(default_voice_to_set) # This trace calls _on_voice_params_changed_for_cache
            self.on_voice_selected(event=None) 
        else:
            self.voice_var.set("") 
            self._update_status(f"语言 '{current_lang_selection}' 没有找到可用语音。")
            self.on_voice_selected(event=None)

    def on_voice_selected(self, event=None):
        selected_voice_name = self.voice_var.get()
        self.role_combo.config(values=["(无)"], state="disabled"); self.role_var.set("(无)") # Triggers _on_voice_params_changed_for_cache
        self.style_combo.config(values=["(默认)"], state="disabled"); self.style_var.set("(默认)") # Triggers _on_voice_params_changed_for_cache
        was_loading_profile = bool(self._profile_being_loaded_settings)
        voices_loaded_ok = isinstance(self.all_voices_in_region, list) and bool(self.all_voices_in_region)
        voice_info_available = selected_voice_name and selected_voice_name in self.current_language_voice_infos

        if not (voices_loaded_ok and voice_info_available):
            status_detail = "" # ... (error status details)
            if not voices_loaded_ok: status_detail = "语音列表未就绪。"
            elif not selected_voice_name: status_detail = "未选择有效语音。"
            elif not voice_info_available: status_detail = f"无法找到语音 '{selected_voice_name}' 的详细信息。"
            final_status = f"语音选择处理失败: {status_detail}"
            if was_loading_profile: final_status = f"加载配置 '{self.profile_var.get()}' 部分失败: {status_detail}"; self._profile_being_loaded_settings = None 
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

        # Role and Style vars are set after this block, their traces will call _on_voice_params_changed_for_cache
        if was_loading_profile and self._profile_being_loaded_settings and \
           not self._profile_being_loaded_settings.get("_voice_application_failed", False) and \
           self._profile_being_loaded_settings["voice"] == selected_voice_name:
            profile_role = self._profile_being_loaded_settings["role"]
            if profile_role in roles_to_display: default_role_to_set = profile_role
            else: messagebox.showwarning("配置加载警告", f"配置角色 '{profile_role}' 对语音 '{selected_voice_name}' 无效。", parent=self.master)
            profile_style = self._profile_being_loaded_settings["style"]
            if profile_style in styles_to_display: default_style_to_set = profile_style
            else: messagebox.showwarning("配置加载警告", f"配置风格 '{profile_style}' 对语音 '{selected_voice_name}' 无效。", parent=self.master)
        if was_loading_profile: self._profile_being_loaded_settings = None 
        
        self.role_var.set(default_role_to_set)
        self.style_var.set(default_style_to_set) 
        
        status_msg = f"语音: {selected_voice_name}"
        if not sdk_roles and not sdk_styles: status_msg += " (无额外角色/风格)"
        if was_loading_profile and not self._profile_being_loaded_settings and self.profile_var.get(): 
             status_msg = f"配置 '{self.profile_var.get()}' 已应用. {status_msg}"
        self._update_status(status_msg)
        self._update_ui_for_playback_state()

    def load_voices_from_azure(self):
        subscription_key = self.subscription_key_entry.get(); service_region = self.service_region_entry.get()
        if not subscription_key or not service_region: messagebox.showerror("配置错误", "请输入有效的 Azure 订阅密钥和区域。", parent=self.master); return
        self._update_status("正在从 Azure 加载语音列表...")
        self.load_voices_button.config(state="disabled")
        self._clear_all_voice_data_and_ui() 
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
                current_lang_after_load = self.language_var.get() 
                if self._pending_profile_to_apply_after_load:
                    p_name = self._pending_profile_to_apply_after_load; self._pending_profile_to_apply_after_load = None 
                    if p_name in self.profile_combo['values']: self.profile_var.set(p_name); self.on_profile_combobox_selected(event=None)
                    else: self._update_status(f"待加载配置 '{p_name}' 未找到。")
                elif current_lang_after_load : self.language_var.set(current_lang_after_load); self.on_language_selected(event=None)
            else: 
                self.all_voices_in_region = []; self.loaded_voices_credentials = {"key": None, "region": None}
                error_msg = f"获取语音列表失败: {result.reason if result else '未知'}" + (f" (详情: {result.cancellation_details.error_details})" if result and result.cancellation_details and result.cancellation_details.error_details else "")
                messagebox.showerror("加载失败", error_msg, parent=self.master); self._update_status(f"加载失败: {error_msg.splitlines()[0]}")
        except Exception as e:
            self.all_voices_in_region = []; self.loaded_voices_credentials = {"key": None, "region": None}
            messagebox.showerror("发生错误", f"加载语音列表时出错: {str(e)}", parent=self.master); self._update_status(f"加载错误: {str(e)}")
        finally: 
            self.load_voices_button.config(state="normal")
            self._update_ui_for_playback_state()

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

    def _build_ssml(self, text_to_speak_raw, lang, voice_name, role, style):
        txt_esc=xml.sax.saxutils.escape(text_to_speak_raw)
        parts=[f'<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xml:lang="{lang}">',f'<voice name="{voice_name}">']
        expr_as=(role!="(无)")or(style!="(默认)");attrs=[]
        if role!="(无)":attrs.append(f'role="{role}"')
        if style!="(默认)":attrs.append(f'style="{style}"')
        if expr_as:parts.extend([f'<mstts:express-as {" ".join(attrs)}>',txt_esc,'</mstts:express-as>'])
        else:parts.append(txt_esc)
        parts.extend(['</voice>','</speak>']);return "".join(parts)

    def _on_closing(self):
        if self.pygame_initialized and pygame.mixer.music.get_busy():
            pygame.mixer.music.stop(); pygame.mixer.music.unload()
        self._cleanup_temp_file()
        if self.pygame_initialized: pygame.mixer.quit(); pygame.quit()
        self.master.destroy()

    def _cleanup_temp_file(self):
        if self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath):
            filepath_to_delete = self.synthesized_audio_filepath
            # 立即清除对文件路径的引用，即使删除失败，下次也不会尝试删除同一个不存在的引用
            self.synthesized_audio_filepath = None 
            
            # 尝试从Pygame卸载当前音乐，以释放文件句柄
            if self.pygame_initialized:
                try:
                    if pygame.mixer.music.get_busy(): # 如果还在播放（理论上不应该，但作为保险）
                        pygame.mixer.music.stop()
                    pygame.mixer.music.unload() # 卸载当前加载的任何音乐
                    # print(f"Debug: Attempted pygame unload for cleanup of {filepath_to_delete}")
                except pygame.error as e:
                    # Pygame可能没有音乐加载，或者mixer未初始化等，这里忽略错误继续尝试删除
                    print(f"Debug: Pygame error during unload in cleanup (ignorable): {e}")
                except Exception as e_pg_unload: # 其他可能的pygame相关错误
                     print(f"Debug: Unexpected error during pygame unload in cleanup: {e_pg_unload}")


            # 尝试删除文件，带延时和重试
            for attempt in range(3): # 最多尝试3次
                try:
                    os.remove(filepath_to_delete)
                    # print(f"Debug: Successfully cleaned up temp file: {filepath_to_delete}")
                    return # 删除成功，退出方法
                except PermissionError as e_perm: # 特别是 WinError 32
                    print(f"Debug: Cleanup attempt {attempt + 1} for '{filepath_to_delete}' failed (PermissionError): {e_perm}")
                    if attempt < 2: # 如果不是最后一次尝试
                        time.sleep(0.1 * (attempt + 1)) # 第一次等0.1s, 第二次等0.2s
                    else:
                        print(f"Debug: Failed to delete temp file after multiple attempts: {filepath_to_delete}")
                except FileNotFoundError:
                    # print(f"Debug: Temp file already deleted or not found during cleanup: {filepath_to_delete}")
                    return # 文件已不存在
                except Exception as e_os_rem: # 其他 os.remove 可能的错误
                    print(f"Debug: Error during os.remove for '{filepath_to_delete}': {e_os_rem}")
                    return # 遇到其他错误，不再重试
        # 如果 self.synthesized_audio_filepath 开始就是 None 或文件不存在，则什么也不做
        self.synthesized_audio_filepath = None # 再次确保清空

    def _format_time(self, seconds):
        if seconds is None or seconds < 0: return "00:00"
        minutes = int(seconds // 60); seconds = int(seconds % 60)
        return f"{minutes:02d}:{seconds:02d}"
    
    def _synthesize_audio_to_file_thread(self):
        current_params = self._get_current_synthesis_params()
        # Use common_inputs for validated Azure key/region for synthesis
        common_inputs = self._get_common_synthesis_inputs(for_playback=True) # This validates all necessary fields
        if not common_inputs:
            self._update_status("输入错误，无法开始合成。")
            self.playback_state = "idle"; self.master.after(0, self._update_ui_for_playback_state); return

        self.playback_state = "synthesizing"
        self.master.after(0, self._update_ui_for_playback_state)
        self._update_status("正在合成语音...")
        
        s_key, s_reg, txt_raw, lang, voice = common_inputs # Use validated inputs
        # Use role and style from current_params as they are part of the synthesis signature
        role, style_val = current_params["role"], current_params["style"]
        
        ssml = self._build_ssml(txt_raw, lang, voice, role, style_val)
        self._cleanup_temp_file()
        try:
            fd, temp_path = tempfile.mkstemp(suffix=".wav", prefix="azure_tts_")
            os.close(fd) 
            self.synthesized_audio_filepath = temp_path
            
            speech_config_obj = speechsdk.SpeechConfig(subscription=s_key, region=s_reg) # Local speech_config
            speech_config_obj.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Riff16Khz16BitMonoPcm)
            audio_config_obj = speechsdk.audio.AudioOutputConfig(filename=self.synthesized_audio_filepath) # Local audio_config
            synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config_obj, audio_config=audio_config_obj)
            result = synthesizer.speak_ssml_async(ssml).get()
            del synthesizer

            if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
                self.total_audio_duration_sec = result.audio_duration.total_seconds() if result.audio_duration else 0
                self.last_synthesis_params = current_params # Cache successful params
                self.text_modified_flag = False 
                self._update_status("合成完毕，准备播放。")
                self.master.after(0, self._start_playback_after_synthesis, True)
            else:
                details = result.cancellation_details if result else None
                msg = f"语音合成取消/失败: {result.reason if result else '未知'}\n" + (f"错误详情: {details.error_details}" if details and details.reason == speechsdk.CancellationReason.Error and details.error_details else "")
                self.master.after(0, lambda m=msg: messagebox.showerror("合成错误", m, parent=self.master))
                self.playback_state = "idle"; self._cleanup_temp_file()
                self.master.after(0, self._update_ui_for_playback_state)
                self.master.after(0, lambda r=result.reason: self._update_status(f"合成错误: {r if r else '未知'}"))
        except Exception as e:
            self.master.after(0, lambda err=str(e): messagebox.showerror("发生严重错误", f"语音合成或文件操作失败: {err}", parent=self.master))
            self.playback_state = "idle"; self._cleanup_temp_file()
            self.master.after(0, self._update_ui_for_playback_state)
            self.master.after(0, lambda err=str(e): self._update_status(f"合成严重错误: {err}"))

    def _start_playback_after_synthesis(self, is_newly_synthesized=False): # is_newly_synthesized not strictly used here anymore but kept for clarity
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

        current_params = self._get_current_synthesis_params() # Get fresh params including text

        if self.playback_state == "playing": 
            if pygame.mixer.music.get_busy(): pygame.mixer.music.pause()
            self.playback_state = "paused"
            if self.progress_updater_id: self.master.after_cancel(self.progress_updater_id); self.progress_updater_id = None
            if self.playback_start_time_monotonic is not None:
                elapsed = time.monotonic() - self.playback_start_time_monotonic
                self.playback_marker_sec += elapsed
                self.playback_start_time_monotonic = None # Paused, so no active start time
            self._update_status("已暂停。")
        
        elif self.playback_state == "paused": 
            pygame.mixer.music.unpause()
            self.playback_state = "playing"
            self.playback_start_time_monotonic = time.monotonic() 
            self._schedule_progress_update()
            self._update_status("继续播放...")
        
        elif self.playback_state == "idle" or self.playback_state == "stopped_by_user": 
            # Check if synthesis is needed
            # Cache is valid if: file exists, text hasn't been flagged as modified by user, AND all other params match
            # Note: _on_voice_params_changed_for_cache clears the file path if voice params changed, forcing re-synthesis
            
            # More robust cache check: compare all relevant current params with last_synthesis_params
            needs_resynthesis = True # Assume re-synthesis is needed by default
            if self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath) and \
               not self.text_modified_flag: # Check text_modified_flag
                # Compare all other relevant parameters
                if all(current_params.get(k) == self.last_synthesis_params.get(k) for k in ["lang", "voice", "role", "style", "subscription_key", "service_region"]) and \
                   current_params["text"] == self.last_synthesis_params.get("text"): # Explicitly check text again
                    needs_resynthesis = False

            if not needs_resynthesis:
                self._update_status("播放已缓存音频...")
                self._start_playback_after_synthesis(is_newly_synthesized=False)
            else: 
                # Check if inputs are valid before starting synthesis thread
                if not self._get_common_synthesis_inputs(for_playback=True):
                     self._update_status("输入不完整，无法播放。")
                     # _get_common_synthesis_inputs already shows a messagebox
                     # Ensure UI reflects that play cannot start
                     self.playback_state = "idle" 
                     self._update_ui_for_playback_state()
                     return

                self._update_status("准备合成新音频...")
                self.total_audio_duration_sec = 0; self.progress_var.set(0) # Reset for new synthesis
                self.time_label_var.set("00:00 / 00:00")
                self.last_synthesis_params = {} # Clear old params as we are re-synthesizing
                
                threading.Thread(target=self._synthesize_audio_to_file_thread, daemon=True).start()
        self._update_ui_for_playback_state()

    def _on_stop_button_click(self):
        if not self.pygame_initialized: return
        
        if self.progress_updater_id: 
            self.master.after_cancel(self.progress_updater_id)
            self.progress_updater_id = None
        
        if pygame.mixer.get_init() and pygame.mixer.music.get_busy(): # 检查mixer是否初始化以及是否有音乐在播放/暂停
            pygame.mixer.music.stop()
            pygame.mixer.music.unload() # <--- 关键：确保在停止后卸载
        
        self.playback_state = "stopped_by_user" 
        self.playback_marker_sec = 0 
        self.playback_start_time_monotonic = None 
        
        self.progress_var.set(0) 
        # 即使停止了，如果之前合成过，总时长信息应该保留以供UI显示
        self.time_label_var.set(f"00:00 / {self._format_time(self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else 0)}")
        
        self._update_status("已停止。")
        self._update_ui_for_playback_state()
        # 注意：这里不再调用 _cleanup_temp_file()。
        # 临时文件的清理主要在以下三种情况发生：
        # 1. _synthesize_audio_to_file_thread() 开始时，清理上一个临时文件。
        # 2. _on_voice_params_changed_for_cache() 或 _on_text_area_modified_flag() 标记缓存失效，间接导致下次播放时清理。
        # 3. _on_closing() 应用退出时。
        # 这样做是为了支持“停止后，如果参数未变，可以重新播放缓存”的功能。


    def _schedule_progress_update(self):
        if self.progress_updater_id: self.master.after_cancel(self.progress_updater_id); self.progress_updater_id = None
        if self.playback_state == "playing" and self.pygame_initialized and pygame.mixer.music.get_busy() and not self.is_user_seeking and self.playback_start_time_monotonic is not None:
            elapsed_time_sec = time.monotonic() - self.playback_start_time_monotonic
            current_display_time_sec = self.playback_marker_sec + elapsed_time_sec
            current_display_time_sec = max(0, min(current_display_time_sec, self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else float('inf')))
            self.progress_var.set(current_display_time_sec)
            self.time_label_var.set(f"{self._format_time(current_display_time_sec)} / {self._format_time(self.total_audio_duration_sec)}")
            if self.total_audio_duration_sec > 0 and current_display_time_sec >= self.total_audio_duration_sec - 0.15: # Adjusted tolerance slightly
                self.master.after(0, self._on_stop_button_click) 
            else: self.progress_updater_id = self.master.after(100, self._schedule_progress_update)
        elif self.playback_state == "playing" and self.pygame_initialized and not pygame.mixer.music.get_busy():
            # If music is not busy but state is playing, it means it finished naturally.
            # Schedule _on_stop_button_click to run in the main Tkinter thread.
            self.master.after(0, self._on_stop_button_click)


    def _on_scale_press(self, event):
        if self.playback_state in ["playing", "paused"] and self.total_audio_duration_sec > 0 and self.pygame_initialized and self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath):
            self.is_user_seeking = True
            if self.playback_state == "playing" and pygame.mixer.music.get_busy(): 
                pygame.mixer.music.pause()
                if self.playback_start_time_monotonic is not None: # Update marker to current pos before drag
                    self.playback_marker_sec += (time.monotonic() - self.playback_start_time_monotonic)
                self.playback_start_time_monotonic = None 

    def _on_scale_release(self, event):
        if self.is_user_seeking and self.pygame_initialized and self.synthesized_audio_filepath and os.path.exists(self.synthesized_audio_filepath):
            self.is_user_seeking = False
            seek_to_sec = self.progress_var.get()
            seek_to_sec = max(0, min(seek_to_sec, self.total_audio_duration_sec if self.total_audio_duration_sec > 0 else 0))

            try:
                # Determine if we should resume playing after seek
                should_be_playing_after_seek = (self.playback_state == "playing") or \
                                               (self.playback_state == "paused" and self.play_pause_button.cget("text") == "▶️ 继续") # Indicates it was playing before drag started

                pygame.mixer.music.stop() # Stop to allow precise start
                # It's crucial that pygame can actually play from a start offset for the format.
                # For WAV, play() then set_pos() is often more reliable than play(start=...).
                pygame.mixer.music.load(self.synthesized_audio_filepath) # Ensure it's loaded
                pygame.mixer.music.play() # Play from beginning
                pygame.mixer.music.set_pos(seek_to_sec) # Then seek to position

                self.playback_marker_sec = seek_to_sec 
                self.playback_start_time_monotonic = time.monotonic()
                
                self.progress_var.set(seek_to_sec) # Ensure scale is at the seeked position
                self.time_label_var.set(f"{self._format_time(seek_to_sec)} / {self._format_time(self.total_audio_duration_sec)}")

                if not should_be_playing_after_seek: # If it was paused (and not playing before drag)
                    pygame.mixer.music.pause()
                    self.playback_state = "paused"
                else: # It was playing or should resume playing
                    self.playback_state = "playing"
                    # If it was paused only for dragging, unpause it (set_pos might leave it paused)
                    if pygame.mixer.music.get_busy() and not self.is_user_seeking : # Check if it actually started
                         pass # Pygame starts playing after set_pos if play() was called
                    else: # If set_pos didn't auto-play or needs unpause
                         pygame.mixer.music.unpause() # Ensure it's unpaused if it should be playing
                    self._schedule_progress_update()
            except Exception as e:
                print(f"Error seeking audio: {e}")
                messagebox.showerror("播放错误", f"音频定位失败: {e}", parent=self.master)
            finally: # Ensure seeking flag is reset and UI updated
                self.is_user_seeking = False
                self._update_ui_for_playback_state()
        elif self.is_user_seeking: # If seeking was true but conditions not met
             self.is_user_seeking = False
             self._update_ui_for_playback_state() # Update UI to reflect non-seeking state

    def _on_scale_drag_changed(self, value_str):
        if self.is_user_seeking and self.pygame_initialized and self.total_audio_duration_sec > 0:
            current_seek_sec = float(value_str)
            current_seek_sec = max(0, min(current_seek_sec, self.total_audio_duration_sec))
            self.time_label_var.set(f"{self._format_time(current_seek_sec)} / {self._format_time(self.total_audio_duration_sec)}")
            # self.progress_var is directly tied to the scale, so it updates automatically.

    # --- Other methods (load_voices_from_azure, save_mp3 etc.) are kept from v4.8.1 ---
    # Ensure they use _update_ui_for_playback_state() where appropriate for button states.
    def save_text_to_mp3_thread(self):
        # Check if save is possible
        inputs = self._get_common_synthesis_inputs()
        if not inputs: # _get_common_synthesis_inputs now shows its own error if validation fails
            self._update_status("输入不完整，无法保存MP3。")
            return # Do not proceed if inputs are not valid

        # Disable all operational buttons during save
        self.play_pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.save_mp3_button.config(state=tk.DISABLED) # Disable itself
        self.progress_bar.config(state=tk.DISABLED)
        self._update_status("准备保存 MP3...")

        # original_play_pause_state and original_stop_state are not strictly needed
        # if _update_ui_for_playback_state handles restoration well.
        threading.Thread(target=self.save_text_to_mp3, daemon=True).start()

    def save_text_to_mp3(self): # Removed original_play_state args
        inputs=self._get_common_synthesis_inputs() # Re-check inputs within thread context if needed, though already checked
        if not inputs: 
            self.master.after(0, lambda: self._update_status("输入错误，MP3保存取消。"))
            self.master.after(0, self._update_ui_for_playback_state) # Restore buttons via main thread
            return
        
        s_key, s_reg, txt_raw, lang, voice = inputs
        role, style_val = self.role_var.get(), self.style_var.get()
        
        # Schedule filedialog in main thread
        filepath = [None] # Use a list to pass by reference from lambda
        def ask_filepath():
            filepath[0] = filedialog.asksaveasfilename(
                defaultextension=".mp3", 
                filetypes=[("MP3 audio file","*.mp3"),("All files","*.*")], 
                title="保存 MP3 文件", 
                initialdir=self.script_dir, 
                parent=self.master
            )
        
        # Run filedialog in main thread and wait if necessary (tricky with threads)
        # Simpler: assume filedialog is okay from a non-main thread if Tkinter is initialized,
        # but it's safer to use self.master.after or a queue if issues arise.
        # For now, direct call as it often works if GUI is stable.
        # If this causes issues, it needs proper main-thread marshalling.
        
        # Let's use a simpler approach for now for filedialog in thread,
        # but be aware it might need self.master.after for robustness on all platforms.
        # For this iteration, assuming direct call works for the user's setup.
        actual_filepath = filedialog.asksaveasfilename(
            defaultextension=".mp3", 
            filetypes=[("MP3 audio file","*.mp3"),("All files","*.*")], 
            title="保存 MP3 文件", 
            initialdir=self.script_dir, 
            parent=self.master # Parent is good
        )
        
        if not actual_filepath: 
            self.master.after(0, lambda: self._update_status("MP3保存已取消"))
            self.master.after(0, self._update_ui_for_playback_state)
            return

        self.master.after(0, lambda p=actual_filepath: self._update_status(f"正在保存到 {os.path.basename(p)}..."))
        ssml = self._build_ssml(txt_raw, lang, voice, role, style_val)
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