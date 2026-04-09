import os
import subprocess

config_json = """{
  "theme": "default",
  "fontFamily": "SimSun, \\"Microsoft YaHei\\", sans-serif"
}"""
with open("config.json", "w", encoding="utf-8") as f:
    f.write(config_json)

puppeteer_config = """{
  "args": ["--no-sandbox"]
}"""
with open("puppeteer-config.json", "w", encoding="utf-8") as f:
    f.write(puppeteer_config)

diagrams = {
    "图4-1_系统总体结构图.mmd": """%%{init: {'theme': 'default', 'themeVariables': { 'fontFamily': 'SimSun, Microsoft YaHei, sans-serif' }}}%%
flowchart TD
    classDef default fill:#FFFFFF,stroke:#000000,stroke-width:1.5px,color:#000000;
    classDef layerFill fill:#FFFFFF,stroke:#000000,stroke-width:1.5px,stroke-dasharray: 5 5,color:#000000,font-weight:bold;

    U1([用户])
    U2([管理员])

    subgraph Web前端["第二层：Web 前端界面"]
        direction LR
        P1[登录页面]
        P2[仪表盘]
        P3[缺陷检测页面]
        P4[历史记录页面]
        P5[日志报表页面]
        P6[销售管理页面]
        P7[缺陷图谱页面]
        P8[用户管理页面]
        P9[系统设置页面]
    end

    subgraph Flask后端["第三层：Flask 后端业务逻辑"]
        direction LR
        M1[用户认证与权限控制模块]
        M2[检测处理模块]
        M3[历史记录管理模块]
        M4[日志报表统计模块]
        M5[销售与库存管理模块]
        M6[系统设置与模型切换模块]
    end

    subgraph 数据与模型["第四层：数据与资源层"]
        direction LR
        D2{{YOLO 缺陷检测模型}}
        D1[(SQLite 数据库)]
        D3[静态资源文件]
        D4[/上传图片目录/]
        D5[/检测结果目录/]
        D6[/缺陷图谱素材/]
    end

    U1 --> Web前端
    U2 --> Web前端
    Web前端 --> Flask后端
    Flask后端 --> D2
    Flask后端 --> D1
    Flask后端 --> D3
    Flask后端 -.-> D4
    Flask后端 -.-> D5
    Flask后端 -.-> D6

    class Web前端,Flask后端,数据与模型 layerFill;
""",

    "图4-2_登录顺序图.mmd": """%%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'edgeLabelBackground':'#ffffff', 'tertiaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
sequenceDiagram
    autonumber
    actor User as 用户
    participant Page as 登录页面
    participant Backend as Flask后端
    participant DB as 数据库

    User->>Page: 输入账号密码并选择角色
    Page->>Backend: 提交登录请求
    Backend->>DB: 查询用户信息
    DB-->>Backend: 返回用户数据
    Backend->>Backend: 校验密码和角色
    
    alt 校验成功
        Backend->>DB: 写入工作日志
        Backend-->>Page: 返回登录成功
        Page-->>User: 跳转到仪表盘
    else 校验失败
        Backend-->>Page: 返回“账号或密码错误”或“身份不匹配”
        Page-->>User: 停留在登录页并显示错误提示
    end
""",

    "图4-3_缺陷检测顺序图.mmd": """%%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'edgeLabelBackground':'#ffffff', 'tertiaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
sequenceDiagram
    autonumber
    actor User as 用户
    participant Page as 检测页面
    participant Backend as Flask后端
    participant YOLO as YOLO模型
    participant DB as 数据库
    participant Inv as 库存表

    User->>Page: 选择图片/拍照
    Page->>Backend: 上传图片
    Backend->>Backend: 保存图片文件
    Backend->>YOLO: 执行缺陷检测
    YOLO-->>Backend: 返回检测结果
    Backend->>Backend: 缺陷映射与等级判定
    Backend->>DB: 写入检测历史
    Backend->>Inv: 更新库存数量
    Backend-->>Page: 返回结果图、缺陷类型、等级、置信度
    Page-->>User: 显示检测结果
""",

    "图4-4_历史与报表顺序图.mmd": """%%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'edgeLabelBackground':'#ffffff', 'tertiaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
sequenceDiagram
    autonumber
    actor User as 用户
    participant Page as 历史/报表页面
    participant Backend as Flask后端
    participant DB as 数据库

    User->>Page: 请求查看数据
    Page->>Backend: 发送查询请求
    Backend->>DB: 查询检测记录和日志数据
    DB-->>Backend: 返回记录数据
    Backend->>Backend: 统计趋势、占比和工作时长
    Backend-->>Page: 返回列表和图表数据
    Page-->>User: 展示历史记录和统计结果
""",

    "图4-5_销售出库顺序图.mmd": """%%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'edgeLabelBackground':'#ffffff', 'tertiaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
sequenceDiagram
    autonumber
    actor User as 用户
    participant Page as 销售页面
    participant Backend as Flask后端
    participant Inv as 库存表
    participant DB as 数据库

    User->>Page: 填写订单信息
    Page->>Backend: 提交销售订单
    Backend->>Inv: 查询等级库存
    Inv-->>Backend: 返回库存数量
    Backend->>Backend: 判断库存是否充足
    
    alt 库存充足
        Backend->>Inv: 扣减库存
        Backend->>DB: 写入销售记录
        Backend-->>Page: 返回创建成功
        Page-->>User: 显示成功提示
    else 库存不足
        Backend-->>Page: 返回库存不足提示
        Page-->>User: 显示失败信息
    end
""",

    "图4-6_主要实体属性图.mmd": """classDiagram
    %%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
    class User {
        +id
        +username
        +password
        +role
    }

    class DetectionHistory {
        +id
        +filename
        +result_image
        +upload_date
        +defect_type
        +confidence
        +user_id
        +grade
    }

    class WorkLog {
        +id
        +user_id
        +login_time
        +logout_time
    }

    class SalesRecord {
        +id
        +customer_name
        +product_grade
        +quantity
        +total_price
        +sale_date
    }

    class SystemSettings {
        +id
        +key_name
        +value
    }

    class Inventory {
        +grade
        +count
    }
""",

    "图4-7_系统E-R图.mmd": """erDiagram
    %%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#ffffff', 'fontFamily': 'SimSun, Microsoft YaHei', 'lineColor': '#000000', 'textColor': '#000000'}}}%%
    User ||--o{ DetectionHistory : "产生检测记录"
    User ||--o{ WorkLog : "产生工作日志"
    Inventory ||--o{ SalesRecord : "支撑销售出库"
    
    User {
    }
    DetectionHistory {
    }
    WorkLog {
    }
    Inventory {
    }
    SalesRecord {
    }
    SystemSettings {
        string desc "独立配置实体"
    }
"""
}

for filename, content in diagrams.items():
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Generated {filename}")

for filename in diagrams.keys():
    png_file = filename.replace('.mmd', '.png')
    cmd = f'npx.cmd -y @mermaid-js/mermaid-cli -i "{filename}" -o "{png_file}" -c config.json -p puppeteer-config.json -b white'
    print(f"Running: {cmd}")
    result = subprocess.run(cmd, shell=True)
    if result.returncode == 0:
        print(f"Successfully generated {png_file}")
    else:
        print(f"Failed to generate {png_file}")
