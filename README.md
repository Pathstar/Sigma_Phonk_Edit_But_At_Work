这是一个可以在任何地方触发Phonk Edit的小程序

Pony Version 为材质替换小马颜艺版

仅可在Windows上运行

bilibili视频: [BV1ckUABwENY](https://www.bilibili.com/video/BV1ckUABwENY)





## Config 介绍

### 主体配置

| 名称                           | 介绍                       | 值               | 默认值           |
| ------------------------------ | -------------------------- | ---------------- | ---------------- |
| is_open_main_window            | 是否打开主窗口             | `true / false`   | true             |
| chance                         | 默认触发概率 (%)           | `0-1`            | 0.4              |
| cooldown                       | 冷却时间 (秒)              | 正数             | 8                |
| volume                         | 音量 (%)                   | `0-1`            | 0.6              |
| min_playtime                   | 最小持续时间 (秒)          | 正数             | 2                |
| max_playtime                   | 最大持续时间 (秒)          | 正数             | 4                |
| min_speed                      | 音频变速最小值             | 正数             | 0.7              |
| max_speed                      | 音频变速最大值             | 正数             | 1.5              |
| texture_scale                  | 贴图大小 (%)               | 正数             | 1                |
| texture_x                      | 贴图在显示器的x位置 (%)    | `0-1`            | 0.5              |
| texture_y                      | 贴图在显示器的y位置 (%)    | `0-1`            | 0.8              |
| is_auto_texture_scale          | 是否启用自动缩放切图       | `true / false`   | true             |
| texture_size                   | 自动缩放宽+高数值 (px)     | 正数             | 500              |
| mouse_triggers_enable          | 是否开启鼠标监测           | `true / false`   | true             |
| mouse_triggers                 | 鼠标配置                   | 见 [鼠标](#鼠标) |                  |
| windows_switch_triggers_enable | 是否开启切换窗口监测       | `true / false`   | true             |
| windows_switch_triggers        | 切换窗口配置               | 见 [窗口](#窗口) |                  |
| background_trigger_rate        | 无条件每秒随机触发概率 (%) | `0-1`            | 0                |
| title                          | 主窗口标题                 | 字符串           | Sigma Phonk Edit |
| subtitle                       | 主窗口副标题               | 字符串           | In Anywhere      |
| dont_do_that                   | just dont...               | `true / false`   | false            |
| is_debug                       | 是否在主窗口打印日志       | `true / false`   | false            |



### 鼠标

- 按键

| 按键   | 介绍     |
| ------ | -------- |
| left   | 左键     |
| right  | 右键     |
| middle | 滚轮按键 |
| x1     | 后侧键   |
| x2     | 前侧键   |

- 配置

| 名称    | 介绍               | 值                            | 默认值 |
| ------- | ------------------ | ----------------------------- | ------ |
| enable  | 是否开启           | `true / false`                | false  |
| press   | 是否按下触发       | `true / false`                | false  |
| release | 是否释放触发       | `true / false`                | true   |
| area    | 触发区域           | `"all" / [[0,0],[1920,1080]]` | all    |
| wait    | 触发后等待激活时间 | 正数 (秒)                     |        |
| chance  | 触发概率           | `"default" / 0-1`             |        |

### 窗口

- 配置

| 名称                    | 介绍                    | 值                   | 默认值          |
| ----------------------- | ----------------------- | -------------------- | --------------- |
| enable                  | 是否开启                | `true / false`       | false           |
| blacklist               | 窗口黑名单 (class_name) | `list['class_name']` | class(任务切换) |
| whitelist               | 窗口白名单 (class_name) | `list['class_name']` |                 |
| windows_detect_interval | 窗口循环监测间隔        | 正数 (秒)            | 0.2             |
| wait                    | 触发后等待激活时间      | 正数 (秒)            | 0               |
| chance                  | 触发概率                | `"default" / 0-1`    | default         |

