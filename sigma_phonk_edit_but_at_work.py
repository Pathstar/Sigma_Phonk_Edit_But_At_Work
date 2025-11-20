import time

start_time_dict = {}
used_time_dict = {}


def print_use_time(time_name, log_prefix=None):
    use_time = time.time() - start_time_dict.get(time_name, 0)
    if log_prefix is None:
        log_prefix = time_name
    print(f"[{log_prefix}] {time_name} Time Used: {use_time}")
    used_time_dict[time_name] = use_time


def record_start_time(time_name):
    start_time_dict[time_name] = time.time()


record_start_time("import")

import os
import json
import random
import sys
import threading
from pynput import mouse
from PIL import ImageTk
import tkinter as tk
import win32gui
import win32api
import mss
from PIL import Image
import soundfile as sf
import sounddevice as sd
import numpy as np

# const attr
SW_PREFIX = "[Main]"
MOUSE_PREFIX = "[Mouse]"
SCREEN_PREFIX = "[Screen]"
SOUND_PREFIX = "[Sound]"
TEXTURE_PREFIX = "[Texture]"




# ---------------- tool ----------------
def load_json(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(e)
        return {}


def get_abs_path(path: str):
    return path
    # return os.path.join(BASE_DIR, path)


def random_chance(chance):
    v = random.random()
    print_xd(f"[Random] random {v} | {chance}")
    return v < chance


def nothing(*args, **kwargs):
    pass


class Config:
    def __init__(self):
        self.mouse_triggers_enable = True
        self.windows_switch_triggers_enable = True
        self.texture_scale = 1
        self.max_playtime = 2
        self.min_playtime = 4
        self.volume = 1
        self.is_debug = False
        self.chance: float = 0.3
        self.cooldown: float = 4
        self.min_speed = 0.7
        self.max_speed = 1.5
        self.mouse_triggers = {}
        self.background_trigger_rate = 0.01
        self.windows_switch_triggers = {}
        self.load_config()

    # override this if config path change
    def load_config(self):
        self.set_config(**load_json(path=get_abs_path('config.json')))
        self.process_config()

    # @staticmethod
    # def config_loader(path='config.json'):
    #     with open(path, 'r', encoding='utf-8') as f:
    #         return json.load(f)

    def set_config(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self, key, value)

    def process_config(self):
        try:
            min_speed = self.min_speed
            max_speed = self.max_speed
            if min_speed > max_speed:
                self.min_speed, self.max_speed = max_speed, min_speed
            # 非法范围保护
            # min_speed = max(0.1, min_speed)  # 不让最小速度过小
            # max_speed = min(3.0, max_speed)  # 最大值人为限定3倍
            # 取随机数并保留1位小数

            min_playtime = self.min_playtime
            max_playtime = self.max_playtime
            if min_playtime > max_playtime:
                self.min_playtime, self.max_playtime = max_playtime, min_playtime

        except Exception as e:
            print(f"config error {e}")
            pass


config: Config = Config()
windows_switch_triggers = config.windows_switch_triggers
windows_blocklist = set(windows_switch_triggers.get("blacklist", {}))
windows_whitelist = set(windows_switch_triggers.get("whitelist", {}))
windows_wait_time = windows_switch_triggers.get("wait", 0)
if windows_wait_time < 0:
    windows_wait_time = 0
windows_detect_interval = windows_switch_triggers.get("windows_detect_interval", 0.2)
if windows_detect_interval <= 0:
    windows_detect_interval = 0.2

windows_chance = windows_switch_triggers.get("chance", "default")
if windows_chance == "default":
    windows_chance = config.chance

print_use_time("import", "Init")
#  pip install pynput pillow pywin32 mss
# pyinstaller --onefile --add-data "config.json;." --add-data "resources;resources" sigma_phonk_edit_but_at_work.py

is_exe = False
if getattr(sys, 'frozen', False):
    # 打包后，sys.frozen会被设置
    BASE_DIR = os.path.dirname(sys.executable)
    is_exe = True
    is_debug = False
else:
    # 本地运行
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if config.is_debug:
    print_xd = print
else:
    print_xd = nothing


# print_xd("test")


def get_focused_monitor_rect():
    hwnd = win32gui.GetForegroundWindow()
    try:
        rect = win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)
    except:
        return None
    cx = (rect[0] + rect[2]) // 2
    cy = (rect[1] + rect[3]) // 2
    # 获取所有显示器及区域
    monitors = win32api.EnumDisplayMonitors()
    for hMonitor, hDC, monitor_rect in monitors:
        l, t, r, b = monitor_rect
        if l <= cx < r and t <= cy < b:
            return (l, t, r, b)
    # 没找到就用主屏
    return win32api.EnumDisplayMonitors()[0][2]


cooldown_dict = {}


def get_cooldown_status(key, cooldown):
    if key in cooldown_dict:
        if time.time() - cooldown_dict[key] < cooldown:
            return False
        else:
            return True
    else:
        return True


def start_cooldown(key):
    cooldown_dict[key] = time.time()


# cooldown attr
SW_COOLDOWN = "SW_CD"


class Playsound:
    def __init__(self, resources=get_abs_path('resources')):
        # self.resources = 'resources'
        self.resources = resources
        self.path = os.path.join(resources, 'sounds')
        self.volumes = load_json(os.path.join(self.resources, 'sounds.json'))
        self.sound_list = self.get_audio_files(self.path)
        self.last_played = ""
        self.last_speed = 1.0

    @staticmethod
    def get_audio_files(folder):
        exts = {'.wav', '.flac', '.ogg'}  # soundfile 不直接支持mp3
        # return [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(exts)]
        sound_files = []
        for root, dirs, files in os.walk(folder):
            for file in files:
                base, ext = os.path.splitext(file)
                if ext.lower() in exts:
                    sound_files.append(os.path.join(root, file))
        return sound_files

    @staticmethod
    def get_audio_duration(file):
        info = sf.info(file)
        duration = info.frames / info.samplerate
        return duration

    @staticmethod
    def modification_speed(data, speed):
        orig_len = len(data)
        new_len = int(orig_len / speed)
        if data.ndim == 1:
            # 单通道
            new_idx = np.linspace(0, orig_len - 1, new_len)
            data_stretched = np.interp(new_idx, np.arange(orig_len), data)
        else:
            # 多通道
            data_stretched = np.stack([
                np.interp(np.linspace(0, orig_len - 1, new_len), np.arange(orig_len), data[:, ch])
                for ch in range(data.shape[1])
            ], axis=1)
        return data_stretched

    def change_speed(self, speed=None):
        if not speed:
            self.last_speed = round(random.uniform(config.min_speed, config.max_speed), 1)
        else:
            self.last_speed = speed

    def ensure_duration(self):
        pass

    # todo ensure_duration

    def random_sound_and_get_duration(self, is_change_speed=True, ensure_duration=False):
        if is_change_speed:
            self.change_speed()

        file = self.get_random_sound()
        self.last_played = file
        file_duration = self.get_audio_duration(file)
        play_duration = file_duration / self.last_speed
        # 如果duration_secs大于音频实际长度，只能播放实际长度
        random_duration = round(random.uniform(config.min_playtime, config.max_playtime), 1)
        play_duration = min(random_duration, play_duration)
        # todo ensure_duration
        return play_duration

    def get_random_sound(self):
        # return r'resources\sounds\phonk7.ogg'
        sound_list = self.sound_list
        if not sound_list:
            print("没有找到音频文件！")
            return ""
        file = random.choice(sound_list)
        if len(sound_list) == 1:
            return file
        while True:
            if file == self.last_played:
                file = random.choice(sound_list)
            else:
                return file

    @staticmethod
    def play_audio_thread(data, samplerate):
        start_time = time.time()
        sd.play(data, samplerate)
        sd.wait()
        print_xd(f"{SOUND_PREFIX} Real time {time.time() - start_time}")

    def play_random_sound(self, duration=4.0, volume=1.0, speed: float = None):
        file = self.last_played
        if not file:
            file = self.get_random_sound()
        else:
            if not os.path.exists(file):
                print("文件不存在")
                return
        self.last_played = file
        basename = os.path.basename(file)

        # 音量平衡
        volumes = self.volumes
        if basename in volumes:
            file_volume = volumes[basename]
            if isinstance(file_volume, (int, float)):
                volume *= file_volume

        # 音量设置
        config_volume = config.volume
        if config_volume > 1:
            config_volume = 1
        volume *= config_volume

        if volume > 1:
            volume = 1

        file_duration = self.get_audio_duration(file)
        speed = self.last_speed if speed is None else speed
        play_duration = round(file_duration / speed, 1)
        play_duration = min(duration, play_duration)
        print_xd(
            f"{SOUND_PREFIX} Playing: {file}, Requested duration: {duration}s, Played duration: {play_duration}s, Speed: {speed}x, Volume: {volume}")
        # 读文件
        data, samplerate = sf.read(file)
        # 裁剪
        end_sample = int(play_duration * speed * samplerate)
        data = data[:end_sample]
        # 变速（变音高）
        if speed != 1.0:
            data = self.modification_speed(data, speed)
            # 更新播放时长：变速会影响实际播放秒数

        # 调整音量
        if volume != 1.0:
            data = data * volume
        # 预防数值溢出
        max_val = np.max(np.abs(data))
        if max_val > 1.0:
            data = data / max_val
        # 播放线程
        t = threading.Thread(target=self.play_audio_thread, args=(data, samplerate))
        t.start()
        # 返回实际播放时间（秒），可以做其他事


# def change_speed(data, speed):
#     '''
#     改变音频速度（和音高） using numpy.resample
#     '''
#     idx = np.arange(0, len(data), 1/speed)
#     idx = idx[idx < len(data)]
#     return data[idx.astype(int)]

button_name_dict = {
    "Button.left": "left",
    "Button.right": "right",
    "Button.middle": "middle",
    "Button.x1": "x1",
    "Button.x2": "x2",
}


class MouseTrigger:
    def __init__(self, **kwargs):
        self.enable = True
        self.press = False
        self.release = True
        self.area: str | list = "all"
        self.screen: str | int = "all"
        self.wait = 0
        self.chance = "default"
        for key, value in kwargs.items():
            setattr(self, key, value)


    def match(self, x, y, pressed):
        if not self.enable:
            return False
        if pressed:
            if not self.press:
                return False
        else:
            if not self.release:
                return False

        area = self.area
        if isinstance(area, list) and len(area) >= 2 and not self.is_point_in_rect(x, y, area[0], area[1]):
            return False

        # chance = self.chance if isinstance(self.chance, (int, float))
        # if isinstance(area, list)

        return True

    @staticmethod
    def is_point_in_rect(x, y, point1, point2):
        # x, y = point[0], point[1]
        x1, y1 = point1[0], point1[1]
        x2, y2 = point2[0], point2[1]

        min_x = min(x1, x2)
        max_x = max(x1, x2)
        min_y = min(y1, y2)
        max_y = max(y1, y2)

        return min_x <= x <= max_x and min_y <= y <= max_y

    def process_config(self):
        try:
            if self.chance == "default":
                self.chance = config.chance
        except Exception as e:
            print(f"config error {e}")
            pass

mouse_trigger_dict = {
    button_name_dict.get(btn, btn): MouseTrigger(**setting)
    for btn, setting in config.mouse_triggers.items()
}


class SigmaWork:
    def __init__(self, resources=get_abs_path('resources')):
        self.activation_count = 0
        self.detection_count = 0
        self.resources = resources
        self.path = os.path.join(self.resources, "textures")
        self.texture_files = self.get_texture_files(self.path)
        self.scales = load_json(os.path.join(self.resources, "textures.json"))
        self.ps = Playsound()
        import queue as pyqueue
        self.tk_queue = pyqueue.Queue()
        self.sigma_work_init()

    def sigma_work_init(self):
        tk_ready = threading.Event()
        tk_thread = threading.Thread(target=self.start_tk_thread, args=(tk_ready, self.tk_queue), daemon=True)
        tk_thread.start()
        tk_ready.wait()

    @staticmethod
    def get_texture_files(path):
        exts = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
        imgs = []
        for root, dirs, files in os.walk(path):
            for file in files:
                base, ext = os.path.splitext(file)
                if ext.lower() in exts:
                    imgs.append(os.path.join(root, file))
        return imgs

    def get_random_texture_image(self):
        return random.choice(self.texture_files)

    # ----------- tool -------------
    @staticmethod
    def scale_image(image, scale):
        # 获取原始尺寸
        width, height = image.size
        # 计算新尺寸
        new_width = int(width * scale)
        new_height = int(height * scale)
        # 使用 LANCZOS 算法进行平滑缩放
        scaled_image = image.resize((new_width, new_height), Image.LANCZOS)
        return scaled_image

    # ----------- Tk管理和多屏展示 -------------
    def start_tk_thread(self, background_ready_event, queue):
        root = tk.Tk()
        root.withdraw()

        def process_queue():
            try:
                while True:
                    func, args, kwargs = queue.get_nowait()
                    func(*args, **kwargs)
            except Exception:
                pass
            root.after(50, process_queue)

        background_ready_event.set()
        process_queue()
        root.mainloop()

    # --------- main logic  ---------
    # --------- 获取当前窗口所在显示器区域 ---------
    def show_bw_screen_for_monitor(self, monitor_rect, duration=2.0):
        img = self.grab_monitor_image(monitor_rect).convert('L').convert('RGB')

        # def do_show(l=monitor_rect[0], t=monitor_rect[1], r=monitor_rect[2], b=monitor_rect[3], img=img, duration=duration):
        def do_show():
            l = monitor_rect[0]
            t = monitor_rect[1]
            r = monitor_rect[2]
            b = monitor_rect[3]
            screen_width = r - l
            screen_height = b - t
            # monitor 1920, 0, 3840, 1080 screen_width 1920 height 1080
            # print(f"monitor {l}, {t}, {r}, {b} screen_width {screen_width} height {screen_height}")
            texture_path = self.get_random_texture_image()
            if texture_path:
                texture = Image.open(texture_path)
                # 缩放处理
                # 获取比较信息
                texture_width, texture_height = texture.size
                # 取短边计算
                if screen_width > screen_height:
                    screen_dim = screen_height
                    texture_dim = texture_height
                    k1_dim = 1080
                else:
                    screen_dim = screen_width
                    texture_dim = texture_width
                    k1_dim = 1920
                # 显示器缩放
                texture_scale = screen_dim / k1_dim

                # 自定义缩放
                texture_scale *= config.texture_scale
                basename = os.path.basename(texture_path)
                if basename in self.scales:
                    scale = self.scales[basename]
                    if isinstance(scale, (int, float)):
                        texture_scale *= scale

                print_xd(f"{TEXTURE_PREFIX} texture_scale {texture_scale}")
                if texture_scale != 1.0:
                    texture = self.scale_image(texture, texture_scale)

                # 计算贴图位置
                center_x = screen_width // 2
                center_y = int(screen_height * 0.8)
                # 贴上去
                # composed = img
                composed = self.overlay_image(img, texture, center_x, center_y, resize_to_bg=True)
            else:
                composed = img

            win = tk.Toplevel()
            win.geometry(f'{screen_width}x{screen_height}+{l}+{t}')
            win.attributes('-topmost', True)
            win.overrideredirect(True)
            edit_bg = ImageTk.PhotoImage(composed)
            label = tk.Label(win, image=edit_bg)
            label.image = edit_bg
            label.pack(fill='both', expand=True)
            # print("wait")
            # time.sleep(1)
            # win.destroy()
            # print("destroy")
            # return
            self.ps.play_random_sound(duration, config.volume)
            win.after(int(duration * 1000), win.destroy)
            # time.sleep(duration)
            # print("destroy")

        self.tk_queue.put((do_show, (), {}))

    def grab_monitor_image(self, monitor_rect):
        l, t, r, b = monitor_rect
        with mss.mss() as sct:
            monitor = {
                "top": t,
                "left": l,
                "width": r - l,
                "height": b - t
            }
            sct_img = sct.grab(monitor)
            img = Image.frombytes('RGB', sct_img.size, sct_img.rgb)
            return img

    def overlay_image(self, bg_img, texture_img, center_x, center_y, resize_to_bg=False):
        """
        在bg_img上，将texture_img的中心贴到(center_x, center_y)。
        bg_img, texture_img均为PIL.Image对象
        resize_to_bg: 若贴图比背景还大，可自动缩小
        返回合成后的新Image
        """
        # bg_w, bg_h = bg_img.size
        tex = texture_img.convert('RGBA')

        # if resize_to_bg:
        #     # 按需缩放大贴图
        #     max_w, max_h = bg_w, bg_h
        #     tw, th = tex.size
        #     scale = min(1.0, min(max_w / tw, max_h / th))
        #     if scale < 1.0:
        #         tex = tex.resize((int(tw*scale), int(th*scale)), Image.ANTIALIAS)

        tw, th = tex.size
        paste_x = int(center_x - tw / 2)
        paste_y = int(center_y - th / 2)
        bg_img = bg_img.convert("RGBA")
        result = bg_img.copy()
        result.paste(tex, (paste_x, paste_y), tex)
        return result

    def entry_sigma(self, wait: float = 0):
        if not get_cooldown_status(SW_COOLDOWN, config.cooldown):
            print_xd(f"{SW_PREFIX} On Cooldown")
            return

        start_cooldown(SW_COOLDOWN)
        trigger_sigma_thread = threading.Thread(target=self.trigger_sigma, args=(wait,), daemon=True)
        trigger_sigma_thread.start()
        # self.trigger_sigma(wait)

    def trigger_sigma(self, wait: float = 0):
        if wait != 0:
            time.sleep(wait)
        monitor_rect = get_focused_monitor_rect()
        if monitor_rect:
            self.activation_count += 1
            print_xd(f"{SW_PREFIX} Active Times: {self.activation_count}")
            self.ps.change_speed()
            self.show_bw_screen_for_monitor(monitor_rect, self.ps.random_sound_and_get_duration())
        else:
            print(f"{SW_PREFIX} Error in getting monitor rect")

    def detected_counter(self):
        print_xd()
        self.detection_count += 1
        print_xd(f"{SW_PREFIX} {time.strftime('%H:%M:%S')} Detection Times: {self.detection_count}")

    def mouse_listener(self):
        def on_click(x, y, btn, pressed):
            # do not use sleep in this scope or mouse will be frozen
            self.detected_counter()
            print_xd(f"{MOUSE_PREFIX} x: {x}, y: {y}, btn: {btn}, pressed: {pressed}")

            mouse_trigger: MouseTrigger = mouse_trigger_dict.get(button_name_dict.get(str(btn)))
            if mouse_trigger is None:
                print_xd(f"{MOUSE_PREFIX} mouse trigger {btn} event is not configured")
                return

            if not mouse_trigger.match(x, y, pressed):
                print_xd(f"{MOUSE_PREFIX} mouse trigger {btn} is configured to be intercepted")
                return

            if not random_chance(mouse_trigger.chance):
                return

            self.entry_sigma(mouse_trigger.wait)

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

    def window_focus_listener(self):
        time.sleep(1)
        last_hwnd = win32gui.GetForegroundWindow()
        time.sleep(0.2)
        while True:
            hwnd = None
            try:
                hwnd = win32gui.GetForegroundWindow()
                if hwnd != last_hwnd and hwnd != 0:
                    last_hwnd = hwnd
                    self.detected_counter()
                    class_name = win32gui.GetClassName(hwnd)
                    if config.is_debug:
                        window_text = win32gui.GetWindowText(hwnd)
                        print_xd(f"{SCREEN_PREFIX} Switched to hwnd={hwnd}, class={class_name}, title={window_text}")
                    #                       任务切换
                    if class_name in windows_blocklist:
                        print_xd(f"{SCREEN_PREFIX} Ignored special window: {class_name}")
                    elif windows_whitelist and class_name not in windows_whitelist:
                        print_xd(f"{SCREEN_PREFIX} Window not in whitelist: {class_name}")
                    else:
                        if random_chance(config.chance):
                            if windows_wait_time != 0:
                                time.sleep(windows_wait_time)
                            self.entry_sigma()
            except Exception as e:
                print(f"{SCREEN_PREFIX} window_focus_listener Failed, hwnd={hwnd}, last_hwnd={last_hwnd}\n{SCREEN_PREFIX} {e}")

            time.sleep(windows_detect_interval)

    def start_sigma_work(self):
        threads = []
        if config.mouse_triggers_enable:
            mouse_thread = threading.Thread(target=self.mouse_listener, daemon=True)
            mouse_thread.start()
            threads.append(mouse_thread)
        if config.windows_switch_triggers_enable:
            window_thread = threading.Thread(target=self.window_focus_listener, daemon=True)
            window_thread.start()
            threads.append(window_thread)
        background_trigger_rate = config.background_trigger_rate
        if background_trigger_rate == 0:
            print("[Background] background_trigger_rate = 0, background trigger is disabled")
            for t in threads:
                t.join()
            return
        while True:
            time.sleep(1)
            random_value = random.random()
            if random_value < background_trigger_rate:
                print_xd(f"[Background] background random triggered: {random_value} < {background_trigger_rate}")
        # 堵塞


if __name__ == '__main__':
    try:
        record_start_time("SW_Init")
        sw = SigmaWork()
        print_use_time("SW_Init")
        time.sleep(2)
        # record_start_time("SW_Start")
        sw.start_sigma_work()
        # print_use_time("SW_Start")
    except Exception as main_error:
        print(main_error)

    # sw.trigger_sigma()
    # time.sleep(3)
    # ps = Playsound()
    # ps.change_speed()
    # ps.play_random_sound(duration=4.5)

    # ps.play_random_sound(duration=5 ,speed=3)

    # time.sleep(4)
