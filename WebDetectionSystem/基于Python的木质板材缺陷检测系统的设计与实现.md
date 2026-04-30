<!-- 来源：C:\Coding\WebDetectionSystem-thesis\WebDetectionSystem\基于Python的木质板材缺陷检测系统的设计与实现.docx -->

目 录

摘要	1

关键词	1

Abstract	2

Keywords	2

1 绪论	5

1.1 研究背景	5

1.2 研究现状	5

1.2.1 国内研究现状	5

1.2.2 国外研究现状	6

1.3 论文组织结构	6

2 关键技术	6

2.1 Python	6

2.2 Flask	7

2.3 YOLO 目标检测算法	7

2.4 SQLite 与 SQLAlchemy	7

3 系统分析	8

3.1 可行性研究	8

3.1.1 经济可行性	8

3.1.2 操作可行性	8

3.1.3 技术可行性	8

3.2 需求分析	8

3.3 系统功能分析	9

4 系统设计	10

4.1 系统结构设计	10

4.2 系统顺序图设计	11

4.2.1 登录顺序图	11

4.2.2 缺陷检测顺序图	12

4.2.3 历史追溯与报表顺序图	13

4.2.4 销售出库顺序图	13

4.3 数据库设计	14

4.3.1 数据库 E-R 图设计	14

4.3.2 数据库表设计	16

5 系统实现	17

5.1 用户功能模块的实现	17

5.1.1 用户登录界面	17

5.1.2 检测历史查询界面	18

5.1.3 日志报表查看界面	19

5.1.4 销售与库存管理界面	21

5.1.5 缺陷图谱与帮助中心界面	21

5.2 管理员功能模块的实现	22

5.2.1 用户管理界面	22

5.2.2 系统设置与模型切换界面	23

5.3 算法检测模块的实现	24

5.3.1 数据获取与预处理	24

5.3.2 检测模型设计与训练	25

5.3.3 系统检测与可视化分析	26

6 系统测试	28

6.1 测试目的	28

6.2 测试环境	28

6.3 测试用例	29

6.3.1 正常功能测试截图	29

6.3.2 异常功能测试截图	30

6.4 测试结论	31

7 总结与展望	31

参考文献	31

致谢	32





基于Python的木质板材表面缺陷系统的设计与实现

计算机科学与技术专业    栾艳康

指导教师    刘荣

摘要：为了克服传统木材表面缺陷检测时人工识别速度慢、检测结果难保存、质量分级和库存管理之间缺少联系等缺点，本文设计并实现了以Python为基础的木材表面缺陷智能检测系统。系统主要使用Python进行开发，用Flask框架做后端业务处理，用SQLite数据库存储用户的个人信息、检测记录、工作日志、库存数据和销售记录，使用YOLO目标检测算法对木材表面缺陷进行识别。系统使用加载好的权重文件对上传的图片进行推理，得到缺陷类别、置信度和检测结果图，最后根据裂纹、死节、活节等缺陷类型进行板材等级判定。系统实现了用户登录、图片上传检测、摄像头采集检测、检测历史查询、日志报表统计、Excel数据导出、库存联动、销售出库、缺陷图谱展示、管理员用户管理、参数设置、模型切换等主要功能。测试结果表明，该系统可以较好地完成木材表面缺陷检测以及业务数据管理工作，有较好的实用价值和教学演示价值。



关键词：Python 木材缺陷检测 YOLO Flask Web 系统 库存销售联动



Design and Implementation of a Python-Based Surface Defect Detection System for Wood Panels

Computer science and technology   Luan Yankang

tutor     Liu Rong

Abstract:In order to solve the issues with manual work on wood surface flaw detection that’s inefficient, hard to record, and has poor connection between grade quality and inventory management, we designed and developed a Python web intelligent detection system for wood surface flaws. System mainly adopts Python as the programming language and uses flask for its backend framework, it is based on sqlite for database operation and uses yolo for object recognition. Use the trained model to do inference on the uploaded wood image, give out categories of defects, their confidence value and result images. And based on the defects found like crack,dead knot and live knot etc. it would go for the timber grading. It has the function of user login, image uploading detection, camera detection, detection history record inquiry, log report statistic output, Excel output, inventory link, outbound call processing, defect knowledge display, administrator user operation, parameter setting and model conversion. From the test results, we can see that this system completes wood surface defect detection and some business data operations, which is also meaningful in terms of teaching demonstrations and graduation projects.



Keywords: Python; Wood Defect Detection; YOLO; Flask; Web System; Inventory and Sales



# 1 绪论

## 1.1 研究背景

随着家具制造、建筑装饰、板材加工等行业的发展，木材质量检测在生产管理中所起的作用也越来越明显。木材表面常见的裂纹、死节、活节等缺陷会严重影响板材的外观、力学性能和后续加工利用率，如果不及时识别并记录下来，很容易造成分级不准、库存统计混乱、销售环节数据失真等问题。传统的检测方式大多依靠人工观察，检测结果受人员认识水平、工作状况及现场环境的影响比较大，不能保证长久的稳定。因此，研究一种可以对木材表面缺陷进行自动识别、结果保存和等级管理的系统，有很强的现实意义。

近些年来，计算机视觉以及深度学习技术在工业检测方面应用十分普遍，目标检测算法可以从图像当中找到目标区域并给出类别信息，给木材缺陷检测赋予了新的技术途径。相比于人工检测方式，基于图像识别的检测方法速度更快，结果可以保存下来，并且方便进行统计分析。YOLO系列算法属于典型的单阶段目标检测算法，在检测速度以及工程部署方面有比较明显的优势，适合用在木材表面缺陷识别系统上。

本文根据以上背景，对木材表面缺陷检测和业务管理需求进行设计并实现一个Web木材表面缺陷智能检测系统。系统不但可以完成缺陷图像的识别，还可以把检测结果同板材等级、检测历史、日志统计、库存变化、销售出库等业务流程结合起来，从而形成一个完整的从图像采集、模型检测、等级判定到数据管理的流程。通过该系统的建立与实现，可以给木材缺陷检测信息化管理提供一种可行的实现方式。

## 1.2 研究现状

### 1.2.1 国内研究现状

国内有关木材缺陷识别、质量分级以及智能制造系统的相关研究也在不断的发展中。近些年来，由于 Python，OpenCV，PyTorch等工具的广泛使用，越来越多的研究把图像处理，机器学习以及深度学习算法应用到木材表面缺陷检测当中。有些研究依靠改良目标检测模型来提升缺陷识别准确率，有的研究则从检测系统创建方面入手，达成图片上传，模型推理以及结果展现等目的。

从总体上看，国内有关研究已经给木材缺陷检测打下了较好的算法基础以及系统的实现思路。但是，在课程设计以及毕业设计类系统中，大部分项目只是实现了上传图片、调用模型、显示检测结果等基本功能，并没有对检测历史保存、角色权限管理、库存联动、销售出库、报表统计等业务环节进行考虑。本文在已有研究的基础上，利用Web系统开发的方法，把木材缺陷检测结果应用到等级判定和业务管理中，使系统更加完整、实用。

### 1.2.2 国外研究现状

国外对于木材缺陷检测以及工业视觉识别的研究开始的比较早。早期的研究主要是利用图像纹理特征、边缘特征、颜色特征对木材表面进行分析，用支持向量机、决策树等传统机器学习的方法对缺陷进行分类。伴随着深度学习技术的发展，卷积神经网络、目标检测网络以及图像分割模型被应用到了工业缺陷检测当中，研究的重点也从单个特征的提取转向了端到端的识别以及实时检测。

国外部分木材加工、金属制造、纺织、电子生产企业已经把机器视觉系统嵌入到生产线上，用以完成质量分级、异常报警、统计分析等工作。有研究显示，缺陷检测系统的作用绝非仅仅在于识别出正确与否的问题，其在检测的稳定程度，设备所具有的适用性，同生产管理系统之间的衔接情况等方面也均具有重要的意义。因此，国外的研究给木材缺陷检测系统的工程化应用提供了很多的参考。

## 1.3 论文组织结构

本文共分为七章，各章内容安排如下。

第1章为绪论，主要介绍课题研究背景、国内外研究现状以及论文整体结构。

第2章为关键技术，主要介绍系统开发过程中使用的 Python、Flask、YOLO 目标检测算法、SQLite 与 SQLAlchemy 等相关技术。

第3章为系统分析，主要从可行性、需求、功能以及性能与安全等方面对系统进行分析。

第4章为系统设计，主要完成系统结构设计、系统顺序图设计和数据库设计。

第5章为系统实现，主要对用户功能模块、管理员功能模块以及算法检测模块进行具体的实现。

第6章为系统测试，主要说明测试目的、测试环境、测试用例和测试结论。第7章为总结与展望，对系统实现成果、存在不足及进一步研究方向进行总结。

# 2 关键技术

## 2.1 Python

Python是解释型高级程序设计语言，具有语法简单、类库丰富、开发效率高这些特点。Python 在 Web 开发、数据处理、人工智能、自动化脚本编写等各方面都有比较广泛的使用。本系统使用 Python 作为主要的开发语言，后端路由处理、数据库操作、YOLO 模型调用、图像文件处理、报表数据导出等都用到了 Python。使用相同的Python开发环境可以减少系统联调的困难程度，并且可以提高开发的速度。

本系统中Python被用作统一的开发语言来完成业务处理和算法调用这两项工作。后端程序使用Python调用Flask框架来完成页面请求响应，使用SQLAlchemy实现数据库操作，使用Ultralytics接口加载目标检测模型进行木材表面缺陷识别。因此Python给Web服务、数据库管理、检测算法的集成提供了一个统一的基础。

## 2.2 Flask

Flask 是一个轻量级的 Web 应用框架，有结构灵活、路由清楚、扩展容易的特点。相比大型Web框架来说，Flask对于项目结构的约束比较小，适合于小型系统以及课程设计项目的快速开发。本系统用 Flask 构建后端服务，用路由函数分别完成登录认证、缺陷检测、历史记录查询、日志报表、销售管理、后台设置等功能。

同时使用Flask-Login实现用户的身份认证以及会话的管理。普通用户和管理员虽然使用同一个系统入口，但是系统会根据用户的角色来控制可以访问的页面以及可以执行的操作，从而保证系统有基本的权限管理功能。

## 2.3 YOLO 目标检测算法

YOLO（You Only Look Once）是一类典型的单阶段目标检测算法，其基本思想是在一次前向推理过程中同时完成目标位置预测和类别判断。相比需要先生成候选区域再分类的两阶段检测方法，YOLO 的检测流程更直接，推理速度较快，适合用于木材表面缺陷这类需要快速定位缺陷区域的图像检测任务。本系统通过 Ultralytics 提供的 YOLO 接口加载训练好的 `best.pt` 权重文件，对上传的木质板材图片进行推理，得到缺陷类别、置信度以及带检测框的结果图。

在系统实现过程中，YOLO 模型输出的原始类别名称并不会直接作为最终业务结果展示。后端程序会先将模型识别到的英文类别映射为中文缺陷名称，再结合缺陷严重程度完成板材等级判定。例如，无明显缺陷时判定为 A 级，检测到活节时判定为 B 级，检测到裂纹或死节时判定为 C 级。通过这种处理方式，模型检测结果能够转化为用户更容易理解的业务信息，也为后续检测历史保存、库存统计和销售出库提供统一的数据基础。

## 2.4 SQLite 与 SQLAlchemy

SQLite 是一种轻量级关系型数据库，具有部署简单、无需独立数据库服务、以文件形式保存数据等特点。对于毕业设计和本地演示类系统来说，SQLite 能够满足基本的数据持久化需求，同时可以降低系统安装、配置和迁移的复杂度。本系统使用 SQLite 保存用户信息、检测历史、工作日志、销售记录、系统参数和库存数据，使检测结果与业务管理数据能够长期保留。

SQLAlchemy 是 Python 中常用的 ORM（Object Relational Mapping，对象关系映射）工具，其作用是将数据库表映射为 Python 类，将表中的记录映射为对象，从而减少手写 SQL 语句的数量。本系统通过 SQLAlchemy 创建用户表、检测历史表、工作日志表、销售记录表、系统设置表和库存表，并在后端业务处理中完成数据查询、筛选、写入和更新操作，为检测记录追溯、报表统计、数据导出、库存联动和销售出库提供数据支持。

# 3 系统分析

## 3.1 可行性研究

### 3.1.1 经济可行性

从经济上来说，本系统主要的投入为开发时间、测试图片准备以及运行环境调试等，不包含高额的商业软件购买费用。项目所使用的 Python、Flask、Ultralytics、OpenCV、Bootstrap 和 Chart.js 等技术组件获取便捷，能够较好地满足毕业设计阶段的开发与演示需求。

### 3.1.2 操作可行性

从系统使用流程上看，系统操作路径比较清楚。登录之后，常用功能的入口就集中出现在左侧的导航栏上，用户可以很快地进入到检测、历史记录、日志报表以及销售管理这些页面当中。检测页面可以采用本地上传或者摄像头采集两种方式，管理员相关的配置都在系统设置页面上完成，可以满足系统日常使用和功能测试的要求。

### 3.1.3 技术可行性

技术可行性分为两方面。一方面，项目所使用的各种技术组件比较成熟，相关资料也比较丰富，有利于开发过程中对问题进行定位和调试。另一方面，系统已经可以完成从登录、检测到库存更新、销售记录写入的全部业务流程，说明现有的技术路线是可行的。

除此之外，系统的核心功能是由后端业务处理模块和图像检测模块来完成的，这样一种结构安排给系统模型替换、业务扩充以及部署改良等提供了一些比较清晰的改进依托。

## 3.2 需求分析

根据系统实际的实现情况，本文把系统使用者分为普通用户和管理员两种角色。普通用户进行登录、缺陷检测、历史记录查看、报表浏览等常规操作，管理员除了以上功能之外还要做用户添加、参数调节、模型切换等工作。这样角色划分和项目中权限控制逻辑保持一致，也为后面的功能分析提供清晰的边界。

系统对检测业务的支持包括单张图片和多张图片上传检测以及浏览器摄像头采集方式，检测完成后返回缺陷类别、置信度、板材等级并保存结果图到服务器指定目录。业务管理中需要对文件名、检测时间、操作人、结果图路径、等级等进行记录并可以按缺陷类型进行检索和分页查询。

系统要完成今日检测量、累计工作时长、缺陷类别占比、近七天检测趋势的汇总，可以导出Excel文件。另外，在检测结束后要对相应等级的库存进行更新，即检测、入库到出库的整个过程，从而形成订单产生、库存减少的过程；使中心能够看到典型的缺陷样张和等级判定依据，方便新手快速掌握操作。

## 3.3 系统功能分析

就系统整体实现方式而言，本系统可以分成页面展示层、业务控制层、模型推理层和数据存储层。分层结构和项目代码组织大体上一致，既可以说明各个模块之间的协作关系，也可以使论文分析的内容和实际系统的实现保持一致。

从角色职责上来说，普通用户主要是使用登录认证、木材缺陷检测、历史记录查询、日志报表查看、销售管理、帮助中心浏览、个人信息维护等功能，管理员在此基础上又增加了用户管理、系统设置、模型切换等后台管理任务。检测模块是系统运行的中心，历史、报表、销售模块一起形成了业务闭环，使中心和后台配置变得简单。

根据页面入口进行细分，系统主要有登录、仪表盘、缺陷检测、历史记录、日志报表、数据导出、销售管理、帮助中心、用户管理、系统设置等九个功能页面。各个页面同后端业务模块一一对应，页面组织关系比较清晰，可以满足日常使用、权限划分以及未来扩展的要求。

就普通用户而言，其主要的操作有登录认证、缺陷检测、历史记录查询、日志报表查看、销售管理、帮助中心浏览、个人中心维护等。普通用户和系统之间功能交互关系图如下图3-1所示。

[图片占位 1]

图3-1 普通用户用例图

从管理员的职责范围来说，在普通用户的权限之上还拥有用户管理、系统设置、模型切换以及日志查看这些管理权限。管理员和系统之间的主要交互关系如图3-2所示。

[图片占位 2]

图3-2 管理员用例图

# 4 系统设计

## 4.1 系统结构设计

该系统采用B/S架构，用户通过浏览器浏览各种功能页面，服务器端做业务处理、模型推理、数据库读写等工作。该架构对于客户端环境要求不高，在本地或者局域网内可以很快地进行部署和演示。系统的总体结构图如下图4-1所示。

从构成上来说，前端页面主要实现图片选择、检测结果展示、图表显示、表单提交的功能；后端服务主要是接收到请求之后执行业务逻辑、调用检测模型返回处理结果；数据库用来保存用户的、检测记录的、工作日志的、库存的、销售的等数据；算法模块主要是对木材表面缺陷进行识别。各个部分一起工作，可以共同支撑系统的主要业务功能。

系统启动时会自动创建上传目录和结果目录，然后进行数据库表的初始化。当系统中没有默认的管理员账号或者初始库存信息时，程序就会写入基础数据，从而保证系统第一次运行之后就可以完成登录、检测以及业务流程演示。

系统用统一布局模板来组织页面。用户登录之后，公共模板会承载起侧边导航、用户信息、消息提示以及功能入口的作用，各个业务页面在统一的框架中会以各自的页面内容来呈现出来。该种设计既保证了界面风格的统一，又利于以后页面的维护以及功能的扩展。

系统业务层主要是由后端路由函数组成的。不同的路由对应着不同的请求，即登录、检测、历史记录、日志报表、销售管理、用户管理、系统设置等请求，在接收到前端的数据之后进行权限判断、数据处理、模型调用、数据库提交等工作。

从数据流的角度来分析，系统主要有检测数据流、统计数据流和销售数据流这三个流程。检测数据流从图片上传开始，经过模型推理和等级判定之后写入检测历史并更新库存，统计数据流从检测记录、工作日志中提取数据生成报表和趋势图，销售数据流根据订单信息完成库存扣减并保存销售记录。三类数据流一起组成了系统的业务链路。

[图片占位 3]

图4-1 系统总体结构图

## 4.2 系统顺序图设计

### 4.2.1 登录顺序图

登录顺序图主要描述的是用户在登录过程中同前端页面、后端服务和数据库之间发生的交互关系。用户在登录页面输入账号、密码并选择身份入口之后，系统先查询用户的账号、密码以及角色是否符合要求，然后进行密码和角色的验证；验证成功之后会将用户的登录状态写入到数据库中，并且记录下用户的日志，如果验证失败，则会给出相应的提示信息。登录顺序图如下图4-2所示。

[图片占位 4]

图4-2 登录顺序图

### 4.2.2 缺陷检测顺序图

缺陷检测顺序图用来说明图片上传、模型推理、结果处理、数据保存之间调用关系。用户提交图片之后，后端先保存原始文件并读取当前检测参数，然后调用YOLO模型进行推理；系统根据模型输出得到检测结果图、缺陷类别、置信度和等级信息，最后将结果写入数据库。缺陷检测顺序图如下图4-3所示。

[图片占位 5]

图4-3 缺陷检测顺序图

### 4.2.3 历史追溯与报表顺序图

历史追溯和报表顺序图主要体现系统对于检测数据进行查询、统计、展示的过程。用户进入历史记录或者日志报表页面之后，系统就会按照当前的身份来确定数据的访问范围，然后依次完成分页查询、缺陷筛选、统计汇总以及图表渲染等工作，以此来实现检测结果的追溯以及运行情况的分析。历史追溯以及报表顺序图如图4-4所示。

[图片占位 6]

图4-4 历史与报表顺序图

### 4.2.4 销售出库顺序图

销售出库顺序图用来描述订单提交和库存扣减之间业务处理的过程。用户在销售页面输入客户名称、板材等级、数量、金额之后，先查询对应等级的库存是否满足出库条件，满足出库条件就执行库存扣减和销售记录保存，否则提示信息并结束操作。销售出库顺序图如下图4-5所示。

[图片占位 7]

图4-5 销售出库顺序图

## 4.3 数据库设计

### 4.3.1 数据库 E-R 图设计

数据库的E-R图用来表示系统中的主要实体以及它们之间的联系。本系统以用户、检测历史、工作日志、销售记录、库存等数据对象为设计对象，用户和检测历史、工作日志之间存在直接关系，库存和销售记录之间是通过板材等级形成的业务关系。主要实体属性图如下图4-6所示。

[图片占位 8]

图4-6 主要实体属性图

根据系统业务流程可知，用户登录之后会产生检测记录和工作日志；检测结果完成等级判定之后会改变库存数量；销售出库操作要读取库存并生成销售记录。系统各个实体之间的联系如图4-7所示。

[图片占位 9]

图4-7 系统 E-R 图

### 4.3.2 数据库表设计

根据系统的功能需求以及E-R图的设计结果，数据库主要是用户表、检测历史表、工作日志表、销售记录表、系统设置表和库存表。各个数据表分别担负着账号管理、检测结果保存、工作时间记载、销售业务存贮、参数设定以及库存统计这些工作。具体表结构设计如上图所示。

表4-1 用户表 User

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| id | Integer | 主键 |
| username | String(150) | 用户名，唯一 |
| password | String(150) | 密码哈希 |
| role | String(20) | 角色，user/admin |



表4-2 检测历史表 DetectionHistory

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| id | Integer | 主键 |
| filename | String(300) | 原始文件名 |
| result_image | String(300) | 结果图文件名 |
| upload_date | DateTime | 上传检测时间 |
| defect_type | String(100) | 缺陷类型 |
| confidence | Float | 最高置信度 |
| user_id | Integer | 关联用户 ID |
| grade | String(10) | 等级结果 |



表4-3 工作日志表 WorkLog

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| id | Integer | 主键 |
| user_id | Integer | 操作用户 ID |
| login_time | DateTime | 登录时间 |
| logout_time | DateTime | 退出时间 |



表4-4 销售记录表 SalesRecord

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| id | Integer | 主键 |
| customer_name | String(100) | 客户名称 |
| product_grade | String(50) | 销售等级 |
| quantity | Integer | 数量 |
| total_price | Float | 订单总价 |
| sale_date | DateTime | 销售时间 |



表4-5 系统设置表 SystemSettings

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| id | Integer | 主键 |
| key_name | String(50) | 设置键名 |
| value | String(200) | 设置值 |



表4-6 库存表 Inventory

| 字段名 | 类型 | 说明 |
| --- | --- | --- |
| grade | String(10) | 等级主键 |
| count | Integer | 库存数量 |



# 5 系统实现

## 5.1 用户功能模块的实现

### 5.1.1 用户登录界面

登陆模块用统一的界面风格设计，普通用户通过不同的标签页进入对应账号入口，管理员通过别的途径进行操作。后端收到登录请求之后，会根据账号信息和入口类型来完成角色匹配校验工作，从而保证不同的身份用户可以按照正确的权限去访问。用户退出系统之后，系统会把工作日志结束时间同步到数据库中，给之后的工作时长统计提供基础数据。

[图片占位 10]

图5-1 登录界面截图

部分关键代码如下：

```python
# app.py
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        login_role = request.form.get('login_role')
        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password, password):
            if user.role != login_role:
                flash('身份不匹配，请切换到对应标签页登录。', 'warning')
                return redirect(url_for('login'))
            login_user(user)
            return redirect(url_for('dashboard'))
    return render_template('login.html')
```

### 5.1.2 检测历史查询界面

历史记录页面的主要工作就是保证检测结果可以被查询、追溯。页面集中显示结果图、文件名、板材等级、缺陷类型、检测时间等信息，可以按缺陷关键词筛选。由于系统运行之后检测记录会不断增多，所以页面使用分页的方式控制单页展示的数量，从而提高查询和浏览的效率。

在删除操作中后端加上了权限控制。管理员可以查看、管理所有的检测数据，普通用户只能查看、操作自己产生的记录。这样处理既可以满足多用户系统权限边界的要求，又可以提高历史数据管理的安全性。

[图片占位 11]

图5-3 历史记录页面截图

部分关键代码如下：

```python
# app.py
@app.route('/history', methods=['GET', 'POST'])
@login_required
def history():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    query = DetectionHistory.query

    if current_user.role != 'admin':
        query = query.filter_by(user_id=current_user.id)
    if search:
        query = query.filter(DetectionHistory.defect_type.contains(search))

    records = query.order_by(
        DetectionHistory.upload_date.desc()
    ).paginate(page=page, per_page=12)
```

### 5.1.3 日志报表查看界面

日志、报表模块主要是对系统中已经保存的检测数据进行统计、整理并做可视化展示。系统可以计算出今天检测量、缺陷类别占比、近七天检测趋势，用前端图表组件展示出来，为系统运行状态分析、检测结果汇总提供支持。

数据导出功能使用pandas、openpyxl对检测记录进行整理，生成Excel文件供浏览器下载。该功能可以把系统内部的结构化数据转变为可以归档、汇总、二次分析的表格文件，从而提高检测结果的数据利用价值。

[图片占位 12]

图5-4 日志与报表页面截图

部分关键代码如下：

```python
# app.py
@app.route('/logs')
@login_required
def logs():
    today_count = hist_query.filter(
        func.date(DetectionHistory.upload_date) == date.today()
    ).count()
    stats_data = [[k, v] for k, v in counter.items()]
    trend_data = []
    for i in range(6, -1, -1):
        target_date = date.today() - timedelta(days=i)
        count = hist_query.filter(
            func.date(DetectionHistory.upload_date) == target_date
        ).count()
        trend_data.append(count)
```

工作时长统计表由 `WorkLog` 表来提供。用户每次登录时系统就会新增一条日志记录，退出时再补写离开时间；报表页面在统计累计工作时长的时候，会同时计算已结束记录和当前仍然处于登录状态的记录，所以页面上显示的是一个随着使用过程不断变化的统计结果。

该统计逻辑虽然实现方式简单，但是使系统记录的内容不再只是检测次数本身，而是将用户的工作时长也纳入进来，从而使得日志数据更加完整，也为报表分析提供更多的参考信息。

[图片占位 13]

图5-5 工作时长统计截图

部分关键代码如下：

```python
# app.py
total_seconds = 0
for l in log_query.filter(WorkLog.logout_time != None).all():
    total_seconds += (l.logout_time - l.login_time).total_seconds()

if session.get('current_log_id'):
    curr = db.session.get(WorkLog, session['current_log_id'])
    if curr:
        total_seconds += (datetime.now() - curr.login_time).total_seconds()

work_time_str = (
    f"{int(total_seconds // 3600)}小时 "
    f"{int((total_seconds % 3600) // 60)}分钟"
)
```

### 5.1.4 销售与库存管理界面

销售和库存管理页面在业务上和检测模块是联动的。页面上显示 A、B、C 三种板材库存数量，下方有销售订单录入表单。用户提交订单之后，系统会对对应的等级库存是否足够进行检查，只有库存满足要求的时候才会执行出库操作并记录销售数据，保证销售数据和库存数据的一致性。

[图片占位 14]

图5-6 销售与库存页面截图

部分关键代码如下：

```python
# app.py
@app.route('/sales', methods=['GET', 'POST'])
@login_required
def sales():
    if request.method == 'POST':
        grade = request.form.get('grade')
        qty = int(request.form.get('quantity'))
        inv = db.session.get(Inventory, grade)
        if not inv or inv.count < qty:
            flash(f'{grade} 库存不足！', 'danger')
        else:
            inv.count -= qty
            sale = SalesRecord(
                customer_name=request.form.get('customer'),
                product_grade=grade,
                quantity=qty
            )
            db.session.add(sale)
```

### 5.1.5 缺陷图谱与帮助中心界面

帮助中心页面给出活节、死节和裂纹三种典型的缺陷样张，给出相应的定义说明、外观特征和等级判定标准。该页面可以辅助用户理解不同的缺陷之间的差别，使检测结果的说明更加直观，也利于新用户快速掌握系统的业务规则。

[图片占位 15]

图5-7 帮助中心页面截图

部分关键代码如下：

```html
# app.py / templates/help.html
@app.route('/help')
@login_required
def help_center():
    return render_template('help.html')

<img src="{{ url_for('static', filename='img/live_knot.png') }}"
     class="card-img-top"
     alt="活节样张">
<img src="{{ url_for('static', filename='img/dead_knot.png') }}"
     class="card-img-top"
     alt="死节样张">
<img src="{{ url_for('static', filename='img/crack.png') }}"
     class="card-img-top"
     alt="裂纹样张">
```

## 5.2 管理员功能模块的实现

### 5.2.1 用户管理界面

用户管理模块只对管理员开放，系统可以新增、删除用户账号，在创建用户的时候会对密码进行哈希处理。本模块实现了系统基础的账号管理功能，为后面权限控制、多用户协同使用打下了基础。

[图片占位 16]

图5-8 用户管理页面截图

部分关键代码如下：

```python
# app.py
@app.route('/admin/users', methods=['GET', 'POST'])
@login_required
def admin_users():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            new_user = User(
                username=request.form.get('username'),
                password=generate_password_hash(request.form.get('password')),
                role=request.form.get('role')
            )
            db.session.add(new_user)
            db.session.commit()
```

### 5.2.2 系统设置与模型切换界面

系统设置页面主要是对检测阈值进行调节，对模型文件进行切换。检测阈值会影响模型推理时目标框保留的条件，模型上传功能使管理员可以在不修改程序代码的情况下替换新的模型权重文件，所以该页面可以较好地体现出系统的可维护性以及扩展性。

[图片占位 17]

图5-9 系统设置页面截图

部分关键代码如下：

```python
# app.py
@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    val = request.form.get('conf_threshold')
    s = SystemSettings.query.filter_by(key_name='conf_threshold').first()
    s.value = val

    f = request.files.get('model_file')
    if f and f.filename.endswith('.pt'):
        fname = secure_filename(f.filename)
        f.save(os.path.join(app.root_path, fname))
        m = SystemSettings.query.filter_by(key_name='current_model').first()
        m.value = fname
```

## 5.3 算法检测模块的实现

### 5.3.1 数据获取与预处理

检测页面属于系统的主要业务页面之一。系统把本地上传、摄像头采集这两种输入方式合并在一个页面上，用选项卡切换检测操作的连续性。页面还加入了本地图片预览的功能，可以方便地让用户在正式提交检测之前检查所选的文件是否正确。

[图片占位 18]

图5-10 检测页面截图

部分关键代码如下：

```html
# templates/detect.html
<ul class="nav nav-tabs mb-3" id="detectTab">
    <li class="nav-item">
        <button class="nav-link active" data-bs-target="#upload-pane"
                type="button" onclick="stopCamera()">本地上传</button>
    </li>
    <li class="nav-item">
        <button class="nav-link" data-bs-target="#camera-pane"
                type="button" onclick="initCamera()">摄像头采集</button>
    </li>
</ul>
<form id="uploadForm">
    <input class="form-control" type="file" id="images"
           name="images" multiple accept="image/*">
    <div class="row g-3 mb-3" id="previewArea"></div>
</form>
```

### 5.3.2 检测模型设计与训练

后端检测流程读取当前的阈值和模型参数，把图片存入到指定的目录里，然后调用相应的推理函数。由于系统模型文件会不断更新，后台设置页面还留有模型切换的功能，所以之后检测任务会优先使用管理员当前所选的模型权重文件。

模型返回的原始类别名称是英文的，系统在结果展示之前会将它转换成中文的板材等级，然后根据已有的业务规则给出板材等级。页面会同时给出缺陷类别以及结果图、置信度、等级判定等信息，使得单次检测可以得到比较完整的业务结果。

[图片占位 19]

图5-11 推理结果展示截图

部分关键代码如下：

```python
# utils/detection.py
def process_image(image_path, save_dir, model_path, conf_threshold=0.25):
    model = load_model(model_path)
    img = cv2.imread(image_path)
    results = model(img, conf=float(conf_threshold))

    detected_objects = []
    max_conf = 0.0
    for r in results:
        res_plotted = r.plot()
        for box in r.boxes:
            cls_id = int(box.cls[0])
            detected_objects.append(model.names[cls_id])
            max_conf = max(max_conf, float(box.conf[0]))
    return {
        'defects': ','.join(list(set(detected_objects))),
        'confidence': round(max_conf, 2)
    }
```

### 5.3.3 系统检测与可视化分析

等级判定不是结果展示阶段附带给出的说明，是检测流程中的重要业务环节。系统默认无缺陷板材为A级，识别到活节为B级，检测到裂纹或死节为C级。等级结果产生之后，后端会同步更新库存表，使单次检测操作可以和库存管理流程自然衔接。

[图片占位 20]

图5-12 等级判定与库存联动截图

部分关键代码如下：

```python
# app.py
detected_grade = 'A级'
if res['is_clean']:
    detected_grade = 'A级'
else:
    for d in res['defects'].split(','):
        zh_name = DEFECT_MAPPING.get(d.strip(), d.strip())
        if zh_name in ['裂纹', '死节']:
            detected_grade = 'C级'
            break
        elif zh_name == '活节':
            detected_grade = 'B级'

inv = db.session.get(Inventory, detected_grade)
if inv:
    inv.count += 1
```

每完成一次检测，系统都会将结果写入 `DetectionHistory` 表。保存的内容有原始文件名、结果图文件名、缺陷类型、最高置信度、操作用户和等级结果等字段；结果图片单独保存到static/results目录下，前端页面再用静态资源路径访问。这样就将图片文件和结构化记录分开存贮，方便页面展示以及报表统计之间的联系。

基于此数据保存方式，历史追溯、图片预览、统计报表等都可以建立在同一批检测数据的基础上，减少重复整理数据的工作量，提高系统各个模块之间数据的一致性。

[图片占位 21]

图5-13 检测结果持久化截图

部分关键代码如下：

```python
# app.py
history = DetectionHistory(
    filename=filename,
    result_image=res['filepath'],
    defect_type=final_defect_str,
    confidence=res.get('confidence', 0.0),
    user_id=current_user.id,
    grade=detected_grade
)
db.session.add(history)
results.append(res)
db.session.commit()
```

# 6 系统测试

## 6.1 测试目的

系统测试的主要目的就是检验木材表面缺陷智能检测系统在目前的运行环境中，是否能够稳定地完成各项主要业务流程，并且检查系统在正常输入和异常输入情况下所做出的反应。本章测试重点为用户登录、缺陷检测、检测记录保存、日志报表统计、Excel数据导出、销售出库和后台设置等，用测试进一步证明系统是否达到了毕业设计阶段的基本使用要求。

## 6.2 测试环境

本次测试环境同系统日常开发环境一致。测试平台使用的是Windows操作系统，后端程序运行在本地Python虚拟环境里，Web框架使用的是Flask，数据库使用的是SQLite，检测模块调用训练好的模型权重文件进行推理；浏览器端主要是对登录、图片上传、摄像头采集、图表展示、销售管理等页面进行测试。

测试时以普通用户和管理员两种不同的身份进入系统，并按照实际的业务流程检验各个主要的功能模块。除了正常的测试之外，对没有选择图片直接提交检测、角色入口和账号身份不匹配、库存不足时继续提交订单等异常情况进行检查，看系统能否给出明确的提示，并且保持稳定运行。

## 6.3 测试用例

| 测试编号 | 测试内容 | 预期结果 | 测试结论 |
| --- | --- | --- | --- |
| T1 | 普通用户登录 | 账号密码正确且角色匹配时成功进入仪表盘 | 通过 |
| T2 | 管理员误用普通用户入口登录 | 系统提示身份不匹配并拒绝登录 | 通过 |
| T3 | 上传单张图片检测 | 返回结果图、缺陷类别、等级和置信度 | 通过 |
| T4 | 批量图片检测 | 多张图片依次返回检测结果并写入历史 | 通过 |
| T5 | 摄像头拍照检测 | 拍照后成功提交并返回检测结果 | 通过 |
| T6 | 按缺陷类型查询历史 | 历史记录支持关键字搜索与分页 | 通过 |
| T7 | 导出 Excel 报表 | 浏览器下载包含检测记录的表格文件 | 通过 |
| T8 | 库存不足时提交销售订单 | 系统阻止出库并提示当前库存 | 通过 |
| T9 | 管理员上传新模型 | 保存成功并在后续检测中优先使用新模型 | 通过 |

表6-1 系统功能测试用例

从测试结果可知，系统在正常使用场景下可以完成登录、检测、记录保存、统计报表、数据导出、库存销售联动等操作，在异常场景下可以对角色不匹配、未选择图片、库存不足等情况进行提示，说明系统已经具有一定的异常处理能力。

### 6.3.1 正常功能测试截图

正常测试时系统可以进行用户登录、图片检测、结果展示和历史记录保存等操作，相关的测试截图如下图6-1所示。

[图片占位 22]

图6-1 正常功能测试截图

### 6.3.2 异常功能测试截图

异常测试时系统会给出未选图片直接提交检测、身份不符登录、库存不足出库等提示，相关的测试截图如下图6-2所示。

[图片占位 23]

图6-2 异常功能测试截图

## 6.4 测试结论

从综合测试结果可知，系统已经可以比较完整的实现论文前文设计的主要功能，即木材表面缺陷图像检测、等级判定、历史追溯、日志报表统计、库存更新、销售出库和管理员后台管理等业务。整体运行比较稳定，可以满足毕业设计阶段对于系统展示和功能测试的基本需求。

但从测试情况看，系统仍有进一步完善空间。历史记录删除之后没有同步清理对应图片文件，数据库目前还是单机部署的SQLite，测试工作也只做功能验证，对高并发访问、异常恢复以及长期运行稳定性等方面的覆盖还比较欠缺。这些问题可以作为后面优化、拓展的出发点。

# 7 总结与展望

本文的研究工作是在已经完成开发和调试的木材表面缺陷检测系统的基础上进行的。论文中大部分页面、业务流程、数据结构都可以在系统实现中找到对应的内容，因此论文分析内容与系统实现情况基本一致。

未来可以从三个方面来开展完善工作。第一，对数据集进行补充，并对模型进行改进，从而提高系统对于复杂木纹、光照变化、多缺陷同时出现等条件的识别稳定性；第二，将现有的单机运行方式逐步拓展为适合多人使用的部署结构，扩大系统的使用范围；第三，继续完善业务层的功能，例如图片文件的清理、权限粒度的控制、日志审计、更完整的销售统计。硬件、时间等条件允许的情况下，系统可以扩大到移动端采集或者现场摄像设备的应用场景。

















# 参考文献

[1]朱咏梅,李玉玲,奚峥皓,等.注意力可变形卷积网络的木质板材瑕疵识别[J].西南大学学报(自然科学版),2024,46(02):159-169.

[2]万怡,贺福强,黄易周,等.多功能木质板材新型干燥工艺研究[J].林业机械与木工设备,2023,51(05):34-39.

[3]胡新悦,杨浩春.定制家居木质板材表面抗油污性能评价方法探讨[J].中国人造板,2024,31(02):35-38.

[4]薛俊峰.S355J0板材表面结疤缺陷原因分析[J].冶金信息导刊,2022,59(02):29-32.

[5]王嘉瑜.常见木质板材燃烧性能研究[J].消防界(电子版),2022,8(15):14-18.

[6]杨艳滨,任国庆,汪超.相控阵超声检测技术在板材中的应用[J].无损检测,2024,46(07):47-53.

[7]杜英国.氧化铝高压溶出过程管道结疤的特性及清理方法优化[J].绿色矿冶,2023,39(03):45-49+67.

[8]吴志轩,黄少楠,黄远民,等.一种基于混合纹理特征的板材表面缺陷检测方法[J].科学技术创新,2024,(16):30-33.

[9]Deqiang Z ,Tianqi J ,Xue G , et al.Investigation on magnetic shield thickness of remote field eddy current probes for inspection of ferromagnetic and non-ferromagnetic plates[J].International Journal of Applied Electromagnetics and Mechanics,2023,71(4):325-339.

[10]Mekhtiyev A ,Sarsikeyev Y ,Gerassimenko T , et al.Development of a magnetic activator to protect an electric water heater against scale formation[J].Eastern-European Journal of Enterprise Technologies,2024,6(1):95-102.

[11]陈宇.基于改进YOLOX算法的木质板材表面缺陷在线检测研究[D].齐鲁工业大学,2024.

[12]何涵.基于改进Yolov5的木质板材封边缺陷检测[D].中国科学院大学(中国科学院西安光学精密机械研究所),2024.

[13]易志浩.基于机器视觉的木质板材表面缺陷检测方法研究[D].佛山科学技术学院,2024.

[14]徐浩,夏振平,林李兴,等.多特征融合的在线板材表面缺陷检测方法研究[J].苏州科技大学学报(自然科学版),2024,41(01):76-84.

[15]凌嘉欣,谢永华.残差神经网络模型在木质板材缺陷分类中的应用[J].东北林业大学学报,2021,49(08):111-116.

# 致谢

本论文的完成，离不开老师的悉心指导、同学的热心帮助以及家人的默默支持。在此，我向所有关心和帮助过我的人表示最衷心的感谢。首先，我诚挚地感谢我的指导教师。在整个毕业设计过程中，老师从课题方向、系统架构、算法设计到论文结构都给予了我细致的指导。在基于Python开发木质板材表面缺陷检测系统时，面对图像预处理、缺陷特征提取、深度学习检测等技术难点，老师总能及时为我指点迷津，使我能够顺利完成系统开发与论文撰写。老师严谨的治学精神与认真负责的态度，让我深受启发，也将成为我未来学习和工作中的宝贵财富。

感谢同窗好友在学习与项目开发中给予我的支持与陪伴，我们相互探讨、共同进步，让毕业设计的过程充实而温暖。感谢家人一直以来的理解、鼓励与支持，是他们的关爱让我能够安心完成学业。感谢所有参与论文评审的专家教授，感谢你们提出的宝贵修改意见，使本论文更加完善。由于本人水平有限，系统与论文中仍存在不足之处，恳请各位老师批评指正。在今后的道路上，我将继续努力学习，不断提升自我，不负师长与家人的期望。



<!-- 共检测到 Word 内嵌图片 25 张，Markdown 中以图片占位标记。 -->
