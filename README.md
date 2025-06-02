# Azure 文本转语音 GUI - 由Gemini 2.5 Pro (05-06) 生成

## 关于

一个带有图形用户界面 (GUI) 的 Python 脚本，它利用您自己的 Azure 语音服务 API 实现文本转语音 (TTS) 功能。此应用程序允许用户轻松配置 Azure 凭据、选择语音、角色和风格，将文本合成为语音进行播放，并将音频保存为 MP3 文件。它还具有语音配置文件管理和音频缓存功能，以提高性能。

## 功能特性

*   **用户友好的 GUI:** 使用 Tkinter 构建，易于交互。
*   **Azure 语音服务集成:** 利用 Azure 实现高质量的文本转语音合成。
*   **语音配置:**
    *   根据您的服务区域直接从 Azure 加载和刷新语音列表。
    *   选择语言、特定语音、角色扮演角色（如果可用）和说话风格（如果可用）。
*   **语音配置文件:**
    *   将当前的语音、角色和风格设置保存为命名配置文件。
    *   快速加载已保存的配置文件。
*   **音频播放:**
    *   集成的音频播放器（使用 Pygame）。
    *   控制：播放、暂停、停止。
    *   带有当前时间/总时长显示的进度条。
    *   通过拖动进度条实现定位功能。
*   **音频缓存:**
    *   合成的音频（WAV 格式）会被临时缓存。
    *   如果文本和语音参数未更改，则会重播缓存的音频，从而节省 API 调用和合成时间。
*   **保存为 MP3:** 直接将语音输出合成并保存到 MP3 文件。
*   **配置持久化:**
    *   Azure 订阅密钥和服务区域保存在本地的 `azure_tts_settings.json` 文件中。
    *   语音配置文件也存储在此 JSON 文件中。
*   **状态更新:** 提供有关操作（如加载语音、合成、保存和错误）的反馈。

## 前置条件

*   Python 3.x
*   一个有效的 Microsoft Azure 帐户。
*   在您的 Azure 门户中创建的 Azure 语音服务资源。您将需要其**订阅密钥**和**服务区域**。

## 安装

1.  **克隆仓库或下载脚本。**
    ```bash
    # 如果您安装了 git
    # git clone https://github.com/Charioteer71/Azure_TTS_GUI.git
    # cd Azure_TTS_GUI
    ```
    或者，直接下载 Python 脚本文件 (`.py`)。

2.  **安装所需的 Python 库:**
    打开您的终端或命令提示符并运行：
    ```bash
    pip install azure-cognitiveservices-speech
    pip install pygame
    ```
    (Tkinter 通常包含在标准 Python 安装中。)

## 配置

1.  **运行脚本:**
    ```bash
    azure_tts_gui_x.x.x(versions).py
    ```

2.  **输入 Azure 凭据:**
    *   在 "Azure 配置" 部分：
        *   **订阅密钥:** 输入您的 Azure 语音服务资源中的密钥。
        *   **服务区域:** 输入您的 Azure 语音服务资源的区域 (例如, `eastus`, `westus2`, `northeurope`)。
    *   点击 **"保存凭据"**。这会将您的密钥和区域保存到与脚本位于同一目录下的 `azure_tts_settings.json` 文件中。

3.  **加载语音:**
    *   点击 **"加载/刷新语音列表"**。**（每次运行都需要点击此按钮）**
    *   这将获取所配置区域的可用语音。将显示成功或错误消息。

## 使用方法

1.  **选择语言:**
    *   语音加载后，从 "选择语言" 下拉菜单中选择一种语言。这将填充 "选择语音" 下拉菜单。

2.  **选择语音:**
    *   从 "选择语音" 下拉菜单中选择一个特定的语音。如果所选语音支持，这可能会填充 "角色风格" 和 "说话风格" 下拉菜单。

3.  **选择角色和风格 (可选):**
    *   如果所选语音可用，请选择 "角色风格" 和/或 "说话风格"。

4.  **输入文本:**
    *   在 "输入文本" 区域键入或粘贴您想转换为语音的文本。

5.  **合成和播放:**
    *   点击 **"▶️ 播放"** 按钮。
        *   如果当前文本和设置的音频尚未合成（或者如果参数已更改），它将首先合成音频（您会看到 "合成中..." 状态），然后播放它。
        *   如果缓存的音频可用且有效，它将立即播放。
    *   使用 **"⏸️ 暂停" / "▶️ 继续"** 和 **"⏹️ 停止"** 按钮进行播放控制。
    *   拖动进度条以在音频中定位。

6.  **保存为 MP3:**
    *   点击 **"保存为 MP3"** 按钮。
    *   将出现一个文件对话框，允许您选择 MP3 文件的位置和名称。

7.  **语音配置文件:**
    *   **保存配置文件:** 配置好所需的语言、语音、角色和风格后，点击 **"保存当前为新配置"**。为配置文件输入一个名称。
    *   **加载配置文件:** 从 "选择配置" 下拉菜单中选择一个已保存的配置文件。应用程序将尝试应用这些设置。如果该区域的语音列表尚未加载，它会提示您先加载。

8.  **正确关闭应用程序:**
    *   为确保临时缓存文件能够被程序正确清理，请务必通过点击应用程序窗口右上角的 **"X" 关闭按钮** 来退出程序。
    *   **避免直接关闭运行此程序的命令提示符（CMD）窗口，** 因为那样会导致程序被强制终止，无法执行正常的清理步骤，可能会留下未删除的临时文件。（程序下次启动时会尝试清理一部分旧的残留文件，但最佳实践是正常关闭GUI窗口。）

## 故障排除

*   **Pygame 初始化失败:** 如果您看到 "Pygame 初始化失败" 错误，播放功能将受限或不可用。请确保 Pygame 已正确安装，并且您的系统具有可用的音频输出。
*   **Azure 凭据错误:** 请仔细检查您的订阅密钥和服务区域。确保语音服务在您的 Azure 门户中处于活动状态。
*   **语音加载失败:** 请检查您的互联网连接以及 Azure 凭据是否正确并具有语音服务的权限。
*   **临时文件清理过程中的 "PermissionError":** 在 Windows 上，如果音频文件仍被 Pygame 锁定，有时会发生这种情况。脚本会尝试重试，但如果问题持续存在，可能表明您的系统上 Pygame 资源释放存在更深层次的问题。

---

# Azure Text-to-Speech GUI - Created by Gemini Pro 2.5 (05-06)

## About

A Python script with a GUI that utilizes your own Azure Speech Service API for Text-to-Speech (TTS) functionality. This application allows users to easily configure Azure credentials, select voices, roles, and styles, synthesize text to speech for playback, and save the audio as an MP3 file. It also features voice profile management and audio caching for improved performance.

## Features

*   **User-Friendly GUI:** Built with Tkinter for easy interaction.
*   **Azure Speech Service Integration:** Leverages Azure for high-quality text-to-speech synthesis.
*   **Voice Configuration:**
    *   Load and refresh voice lists directly from Azure based on your service region.
    *   Select language, specific voice, role-play character (if available), and speaking style (if available).
*   **Voice Profiles:**
    *   Save current voice, role, and style settings as named profiles.
    *   Quickly load saved profiles.
*   **Audio Playback:**
    *   Integrated audio player (uses Pygame).
    *   Controls: Play, Pause, Stop.
    *   Progress bar with current time/total duration display.
    *   Seek functionality by dragging the progress bar.
*   **Audio Caching:**
    *   Synthesized audio (WAV format) is temporarily cached.
    *   If the text and voice parameters haven't changed, the cached audio is replayed, saving API calls and synthesis time.
*   **Save as MP3:** Synthesize and save the speech output directly to an MP3 file.
*   **Configuration Persistence:**
    *   Azure subscription key and service region are saved locally in `azure_tts_settings.json`.
    *   Voice profiles are also stored in this JSON file.
*   **Status Updates:** Provides feedback on operations like loading voices, synthesizing, saving, and errors.

## Prerequisites

*   Python 3.x
*   An active Microsoft Azure account.
*   An Azure Speech Service resource created in your Azure portal. You will need its **Subscription Key** and **Service Region**.

## Installation

1.  **Clone the repository or download the script.**
    ```bash
    # If you have git installed
    # git clone https://github.com/Charioteer71/Azure_TTS_GUI.git
    # cd Azure_TTS_GUI
    ```
    Otherwise, just download the Python script file (`.py`).

2.  **Install required Python libraries:**
    Open your terminal or command prompt and run:
    ```bash
    pip install azure-cognitiveservices-speech
    pip install pygame
    ```
    (Tkinter is usually included with standard Python installations.)

## Configuration

1.  **Run the script:**
    ```bash
    azure_tts_gui_x.x.x(versions).py
    ```

2.  **Enter Azure Credentials:**
    *   In the "Azure 配置" (Azure Configuration) section:
        *   **订阅密钥 (Subscription Key):** Enter the key from your Azure Speech Service resource.
        *   **服务区域 (Service Region):** Enter the region for your Azure Speech Service resource (e.g., `eastus`, `westus2`, `northeurope`).
    *   Click **"保存凭据" (Save Credentials)**. This will save your key and region to `azure_tts_settings.json` in the same directory as the script.

3.  **Load Voices:**
    *   Click **"加载/刷新语音列表" (Load/Refresh Voice List)**. **(Click this button every time you run it.)**
    *   This will fetch available voices for the configured region. A success or error message will be displayed.

## Usage

1.  **Select Language:**
    *   Once voices are loaded, choose a language from the "选择语言" (Select Language) dropdown. This will populate the "选择语音" (Select Voice) dropdown.

2.  **Select Voice:**
    *   Choose a specific voice from the "选择语音" (Select Voice) dropdown. This may populate the "角色风格" (Role) and "说话风格" (Style) dropdowns if the selected voice supports them.

3.  **Select Role and Style (Optional):**
    *   If available for the chosen voice, select a "角色风格" (Role) and/or "说话风格" (Style).

4.  **Enter Text:**
    *   Type or paste the text you want to convert to speech into the "输入文本" (Input Text) area.

5.  **Synthesize and Play:**
    *   Click the **"▶️ 播放" (Play)** button.
        *   If the audio for the current text and settings hasn't been synthesized yet (or if parameters changed), it will first synthesize the audio (you'll see a "合成中..." status) and then play it.
        *   If cached audio is available and valid, it will play immediately.
    *   Use the **"⏸️ 暂停" (Pause) / "▶️ 继续" (Resume)** and **"⏹️ 停止" (Stop)** buttons for playback control.
    *   Drag the progress bar to seek through the audio.

6.  **Save as MP3:**
    *   Click the **"保存为 MP3" (Save as MP3)** button.
    *   A file dialog will appear, allowing you to choose the location and name for your MP3 file.

7.  **Voice Profiles:**
    *   **Save Profile:** After configuring your desired Language, Voice, Role, and Style, click **"保存当前为新配置" (Save Current as New Profile)**. Enter a name for the profile.
    *   **Load Profile:** Select a saved profile from the "选择配置" (Select Profile) dropdown. The application will attempt to apply the settings. If the voice list for the region hasn't been loaded, it will prompt you to load it first.

8.  **Closing the Application Correctly:**
    *   To ensure that temporary cache files are properly cleaned up by the program, always exit the application by clicking the **"X" close button** on the application window's title bar.
    *   **Avoid directly closing the Command Prompt (CMD) window** that might be running this program. Doing so will forcibly terminate the application, preventing it from performing its normal cleanup procedures, which may leave temporary files undeleted. (The application will attempt to clean up some old orphaned files on its next startup, but the best practice is to close the GUI window normally.)
## Troubleshooting

*   **Pygame Initialization Error:** If you see a "Pygame 初始化失败" (Pygame Initialization Failed) error, playback functionality will be limited or unavailable. Ensure Pygame is correctly installed and your system has a working audio output.
*   **Azure Credentials Error:** Double-check your Subscription Key and Service Region. Ensure the Speech Service is active in your Azure portal.
*   **Voice Loading Fails:** Verify your internet connection and that the Azure credentials are correct and have permissions for the Speech Service.
*   **"PermissionError" during temp file cleanup:** This can sometimes happen on Windows if the audio file is still locked by Pygame. The script attempts retries, but if persistent, it might indicate a deeper issue with Pygame's resource release on your system.
