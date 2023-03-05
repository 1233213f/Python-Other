## 批量下载Altair网站上的公开报告，节约下载时间，快速查阅资料
## 2019-08-23
## code by Tecson


import requests

json ={
    "data": {
        "attachs": [{
            "path": "doc/20190727011002927-工程大数据、云计算、机器学习与HPC分会场 凌巍才  Dell EMC HPC Solution.pdf",
            "intime": "2019-07-27 06:09:04",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "工程大数据、云计算、机器学习与HPC分会场 凌巍才  Dell EMC HPC Solution.pdf",
            "sort": 2655,
            "attachid": "atalx3az8xgv8n1g0wbbqevnerfgg4v2",
            "type": "3"
        }, {
            "path": "doc/20190720015039871-工程大数据、云计算、机器学习与HPC分会场 邢志远 SEM虚拟化车辆性能平台的建设与应用.pdf",
            "intime": "2019-07-20 06:49:41",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "工程大数据、云计算、机器学习与HPC分会场 邢志远 SEM虚拟化车辆性能平台的建设与应用.pdf",
            "sort": 2657,
            "attachid": "at3ygx4g1411tiqn75pgbwb5vdtks5u4",
            "type": "3"
        }, {
            "path": "doc/20190720015039107-工程大数据、云计算、机器学习与HPC分会场 王轶华 PBS Works加速人工智能创新平台.pdf",
            "intime": "2019-07-20 06:49:40",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "工程大数据、云计算、机器学习与HPC分会场 王轶华 PBS Works加速人工智能创新平台.pdf",
            "sort": 2658,
            "attachid": "ateg7dyc8yer2ypmvlfymtbv54g9ea9t",
            "type": "3"
        }, {
            "path": "doc/20190720015037854-工程大数据、云计算、机器学习与HPC分会场 庞成晨 奥工混合态异构超算云解决方案.pdf",
            "intime": "2019-07-20 06:49:39",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "工程大数据、云计算、机器学习与HPC分会场 庞成晨 奥工混合态异构超算云解决方案.pdf",
            "sort": 2656,
            "attachid": "at4pnw4p2zgaswio2zc5plqgshyumg1i",
            "type": "3"
        }, {
            "path": "doc/20190719231833439-NVH全球最佳解决方案分会场 管建民 Driving More Physics into NVH Simulations.pdf",
            "intime": "2019-07-20 04:17:34",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "NVH全球最佳解决方案分会场 管建民 Driving More Physics into NVH Simulations.pdf",
            "sort": 2678,
            "attachid": "at2o2la427l815npnvquxln7uglys0fq",
            "type": "3"
        }, {
            "path": "doc/20190719231832864-NVH全球最佳解决方案分会场 张森 基于MSA方法的方向盘摆振整车仿真分析.pdf",
            "intime": "2019-07-20 04:17:34",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "NVH全球最佳解决方案分会场 张森 基于MSA方法的方向盘摆振整车仿真分析.pdf",
            "sort": 2677,
            "attachid": "atb7ll69cnt8zxsuc1npy8yx5gpnizb6",
            "type": "3"
        }, {
            "path": "doc/20190719231824864-NVH全球最佳解决方案分会场 张冰冰 Driveline NVH Simulation with Altair Tool Chain.pdf",
            "intime": "2019-07-20 04:17:26",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "NVH全球最佳解决方案分会场 张冰冰 Driveline NVH Simulation with Altair Tool Chain.pdf",
            "sort": 2676,
            "attachid": "at0gz0v9b2k5m1c2r0dy4c1ghtr3sgkm",
            "type": "3"
        }, {
            "path": "doc/20190719231821855-NVH全球最佳解决方案分会场 李金录 基于NVHD的动力总成激励车内轰鸣声解决方案.pdf",
            "intime": "2019-07-20 04:17:23",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "NVH全球最佳解决方案分会场 李金录 基于NVHD的动力总成激励车内轰鸣声解决方案.pdf",
            "sort": 2675,
            "attachid": "atygxyhc3x9gonpbdi22ku8skz7s3c5z",
            "type": "3"
        }, {
            "path": "doc/20190719231801859-3D工业设计与制造仿真分会场 漆伟 仿真驱动设计和制造.pdf",
            "intime": "2019-07-20 04:17:03",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 漆伟 仿真驱动设计和制造.pdf",
            "sort": 2674,
            "attachid": "at4kliz9rucuz2pkas2np30ofedfd59g",
            "type": "3"
        }, {
            "path": "doc/20190719231800864-3D工业设计与制造仿真分会场 李俊 多物理场仿真一站式解决方案0718.pdf",
            "intime": "2019-07-20 04:17:02",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 李俊 多物理场仿真一站式解决方案0718.pdf",
            "sort": 2673,
            "attachid": "atogy4ps6kixggzxkezm5x2o945fthqn",
            "type": "3"
        }, {
            "path": "doc/20190719231759887-3D工业设计与制造仿真分会场 陈伟琦 工业造型设计与效果表现方法.pdf",
            "intime": "2019-07-20 04:17:01",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 陈伟琦 工业造型设计与效果表现方法.pdf",
            "sort": 2672,
            "attachid": "atybm28kizgr16gv8ezhtugiywp48zax",
            "type": "3"
        }, {
            "path": "doc/20190719231754864-3D工业设计与制造仿真分会场 徐成斌 设计优化、工艺仿真与3D打印设计.pdf",
            "intime": "2019-07-20 04:16:56",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 徐成斌 设计优化、工艺仿真与3D打印设计.pdf",
            "sort": 2671,
            "attachid": "atubbqfk8k08veyqp2voq10ev9bkdzlk",
            "type": "3"
        }, {
            "path": "doc/20190719231750863-3D工业设计与制造仿真分会场 裘旭东 solidThinking在民机中的应用.pdf",
            "intime": "2019-07-20 04:16:52",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 裘旭东 solidThinking在民机中的应用.pdf",
            "sort": 2669,
            "attachid": "atx264kvngkspigqot7155brxbtcfrrx",
            "type": "3"
        }, {
            "path": "doc/20190719231747860-3D工业设计与制造仿真分会场 夏成兴 模流分析指导模具开发前到模具设计后的应用.pdf",
            "intime": "2019-07-20 04:16:49",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "3D工业设计与制造仿真分会场 夏成兴 模流分析指导模具开发前到模具设计后的应用.pdf",
            "sort": 2670,
            "attachid": "atyz89lvz4y60o0egzov9bfgno4ylmzk",
            "type": "3"
        }, {
            "path": "doc/20190719231725878-结构及多学科仿真分会场 基于虚拟路面的车辆动力学CAE解决方案.pdf",
            "intime": "2019-07-20 04:16:27",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "结构及多学科仿真分会场 基于虚拟路面的车辆动力学CAE解决方案.pdf",
            "sort": 2668,
            "attachid": "atb4ersu4qo1i7cx64w74k6fl5dsg7b3",
            "type": "3"
        }, {
            "path": "doc/20190719231723855-结构及多学科仿真分会场 许光浩 Altair多学科解决方案及应用.pdf",
            "intime": "2019-07-20 04:16:25",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "结构及多学科仿真分会场 许光浩 Altair多学科解决方案及应用.pdf",
            "sort": 2667,
            "attachid": "attf5atb63glzu5k6lwymvnn8v4qxpg4",
            "type": "3"
        }, {
            "path": "doc/20190719231722862-结构及多学科仿真分会场 Michael Hoffmann _Innovating Mechatronics Products.pdf",
            "intime": "2019-07-20 04:16:24",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "结构及多学科仿真分会场 Michael Hoffmann _Innovating Mechatronics Products.pdf",
            "sort": 2666,
            "attachid": "atqxwoo81gfpgcoiv86p6fcdeybsfhtr",
            "type": "3"
        }, {
            "path": "doc/20190719231718861-结构及多学科仿真分会场 汤凯利 SimSolid_无网格快速仿真分析工具.pdf",
            "intime": "2019-07-20 04:16:20",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "结构及多学科仿真分会场 汤凯利 SimSolid_无网格快速仿真分析工具.pdf",
            "sort": 2665,
            "attachid": "atubtw4zisbeg3ai4gimg6gruyoo59h0",
            "type": "3"
        }, {
            "path": "doc/20190719231715860-结构及多学科仿真分会场 刘勇 HyperWorks结构求解器最新进展.pdf",
            "intime": "2019-07-20 04:16:17",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "结构及多学科仿真分会场 刘勇 HyperWorks结构求解器最新进展.pdf",
            "sort": 2664,
            "attachid": "at8rkryv851o3gx6c4z6te9x9ofm6wd5",
            "type": "3"
        }, {
            "path": "doc/20190719231557864-建模可视化及流程自动化分会场 熊春明 SimLab最新进展和未来演进.pdf",
            "intime": "2019-07-20 04:14:59",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "建模可视化及流程自动化分会场 熊春明 SimLab最新进展和未来演进.pdf",
            "sort": 2663,
            "attachid": "at7bdelmvvc2hgdk6fu7116u1dv36d1f",
            "type": "3"
        }, {
            "path": "doc/20190719231556855-建模可视化及流程自动化分会场 张晨 HyperMesh新界面及新功能演示.pdf",
            "intime": "2019-07-20 04:14:58",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "建模可视化及流程自动化分会场 张晨 HyperMesh新界面及新功能演示.pdf",
            "sort": 2662,
            "attachid": "atqg8egughhx9zq8ydrboinf0567k66r",
            "type": "3"
        }, {
            "path": "doc/20190719231554855-建模可视化及流程自动化分会场 吴莉洁 Altair数值计算工具Compose在CAE和数据处理中的应用演示.pdf",
            "intime": "2019-07-20 04:14:56",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "建模可视化及流程自动化分会场 吴莉洁 Altair数值计算工具Compose在CAE和数据处理中的应用演示.pdf",
            "sort": 2661,
            "attachid": "atlwwu3pwkbs79b47a3h3qyio8eixek7",
            "type": "3"
        }, {
            "path": "doc/20190719231551854-建模可视化及流程自动化分会场 张发 盒段级航空复合材料结构精细有限元快速建模技术.pdf",
            "intime": "2019-07-20 04:14:53",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "建模可视化及流程自动化分会场 张发 盒段级航空复合材料结构精细有限元快速建模技术.pdf",
            "sort": 2660,
            "attachid": "atqm7pcc8fg25c9pynb8coi4tnip4ebm",
            "type": "3"
        }, {
            "path": "doc/20190719231544864-建模可视化及流程自动化分会场 肖守春 AI仿真在通讯产品中的应用.pdf",
            "intime": "2019-07-20 04:14:46",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "建模可视化及流程自动化分会场 肖守春 AI仿真在通讯产品中的应用.pdf",
            "sort": 2659,
            "attachid": "ati75wm458kda8az4a81y3b9p0hi1qwr",
            "type": "3"
        }, {
            "path": "doc/20190719231515858-高频电磁仿真及智联网分会场 焦金龙 天线罩电磁热一体化仿真.pdf",
            "intime": "2019-07-20 04:14:17",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 焦金龙 天线罩电磁热一体化仿真.pdf",
            "sort": 2654,
            "attachid": "atd09wafflut7ylkoh1egxuetksqslrm",
            "type": "3"
        }, {
            "path": "doc/20190719231509857-高频电磁仿真及智联网分会场 方重华 船舶电磁兼容与电磁隐身性能预测.pdf",
            "intime": "2019-07-20 04:14:11",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 方重华 船舶电磁兼容与电磁隐身性能预测.pdf",
            "sort": 2653,
            "attachid": "atc9gv9msq396zxgq4xpigug6gu32bmg",
            "type": "3"
        }, {
            "path": "doc/20190719231509372-高频电磁仿真及智联网分会场 王晓峰 Altair Feko+Winprop 2019 新功能及技术亮点.pdf",
            "intime": "2019-07-20 04:14:10",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 王晓峰 Altair Feko+Winprop 2019 新功能及技术亮点.pdf",
            "sort": 2652,
            "attachid": "at027hh0sb5y21v5psxea3y38glp366q",
            "type": "3"
        }, {
            "path": "doc/20190719231506862-高频电磁仿真及智联网分会场 贺苒 Application of Altair Feko simulation in vehicle EMC.pdf",
            "intime": "2019-07-20 04:14:08",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 贺苒 Application of Altair Feko simulation in vehicle EMC.pdf",
            "sort": 2651,
            "attachid": "atl23m5f5xroep34tz69t0tdwsgyr19t",
            "type": "3"
        }, {
            "path": "doc/20190719231505868-高频电磁仿真及智联网分会场 张小林 Altair Feko在相控阵天线仿真中的应用.pdf",
            "intime": "2019-07-20 04:14:07",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 张小林 Altair Feko在相控阵天线仿真中的应用.pdf",
            "sort": 2650,
            "attachid": "at0g3g9o4ms6ptrwgqp3rgvi8wm82igt",
            "type": "3"
        }, {
            "path": "doc/20190719231458858-高频电磁仿真及智联网分会场 高伟 电磁散射的高效建模与仿真应用.pdf",
            "intime": "2019-07-20 04:14:00",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "高频电磁仿真及智联网分会场 高伟 电磁散射的高效建模与仿真应用.pdf",
            "sort": 2649,
            "attachid": "atsmffznssn9ehrwc18zsdlowdk2m8ge",
            "type": "3"
        }, {
            "path": "doc/20190719231440872-CFD技术分会场 Frank Wu 无需妥协，达成气动性能设计目标.pdf",
            "intime": "2019-07-20 04:13:42",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "CFD技术分会场 Frank Wu 无需妥协，达成气动性能设计目标.pdf",
            "sort": 2648,
            "attachid": "atfvw1ovscg2xt0c9o9gibwpggmnkmk0",
            "type": "3"
        }, {
            "path": "doc/20190719231416855-CFD技术分会场 武珑 Altair VWT整车外气动性能解决方案.pdf",
            "intime": "2019-07-20 04:13:18",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "CFD技术分会场 武珑 Altair VWT整车外气动性能解决方案.pdf",
            "sort": 2647,
            "attachid": "atfg2bhucdy9q2dispwws9u8ve6aq7ta",
            "type": "3"
        }, {
            "path": "doc/20190719231415857-CFD技术分会场 陈刚 nanoFluidX在传动系统的润滑与冷却分析应用.pdf",
            "intime": "2019-07-20 04:13:17",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "CFD技术分会场 陈刚 nanoFluidX在传动系统的润滑与冷却分析应用.pdf",
            "sort": 2646,
            "attachid": "atygwesn6h50mhqtfgat6md2bn2qeu29",
            "type": "3"
        }, {
            "path": "doc/20190719231408869-CFD技术分会场 王轶华 GPU高性能计算及HPC平台优化.pdf",
            "intime": "2019-07-20 04:13:10",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "CFD技术分会场 王轶华 GPU高性能计算及HPC平台优化.pdf",
            "sort": 2645,
            "attachid": "at9o67fnphyvbpv89vt19in93lz9fcoe",
            "type": "3"
        }, {
            "path": "doc/20190719231406862-CFD技术分会场 王翔 基于Co-simulation技术的洗衣机平衡圈优化设计.pdf",
            "intime": "2019-07-20 04:13:08",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "CFD技术分会场 王翔 基于Co-simulation技术的洗衣机平衡圈优化设计.pdf",
            "sort": 2644,
            "attachid": "atbr9nzvas2efsb4niuy7b9hqi9kabfr",
            "type": "3"
        }, {
            "path": "doc/20190719231331856-显式非线性技术分会场 Jean Michel 面向工程应用的Radioss_2019.pdf",
            "intime": "2019-07-20 04:12:33",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "显式非线性技术分会场 Jean Michel 面向工程应用的Radioss_2019.pdf",
            "sort": 2643,
            "attachid": "atnumk9yob0qhzgv8za1kd245zcifhdw",
            "type": "3"
        }, {
            "path": "doc/20190719231329865-显式非线性技术分会场 连昌伟 碰撞仿真材料卡的建立及Altair与宝钢材料库合作.pdf",
            "intime": "2019-07-20 04:12:31",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "显式非线性技术分会场 连昌伟 碰撞仿真材料卡的建立及Altair与宝钢材料库合作.pdf",
            "sort": 2642,
            "attachid": "ati3hl7wuumy79bfs1zvoee3qqg9qmwx",
            "type": "3"
        }, {
            "path": "doc/20190719231327858-显式非线性技术分会场 张伟伟 Radioss_AMS方法及其应用.pdf",
            "intime": "2019-07-20 04:12:29",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "显式非线性技术分会场 张伟伟 Radioss_AMS方法及其应用.pdf",
            "sort": 2641,
            "attachid": "atxuzcr1i15cgoigp7ah1cqssx8mra2o",
            "type": "3"
        }, {
            "path": "doc/20190719231249857-轻量化技术分会场 徐自立 Altair最新优化技术.pdf",
            "intime": "2019-07-20 04:11:51",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "轻量化技术分会场 徐自立 Altair最新优化技术.pdf",
            "sort": 2640,
            "attachid": "atkwoxn1doqp6ra2ydfmigwmx8ul0l1v",
            "type": "3"
        }, {
            "path": "doc/20190719231240859-轻量化技术分会场 卢顺杰 试验协同仿真- VIC-3D非接触全场应变测量系统.pdf",
            "intime": "2019-07-20 04:11:42",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "轻量化技术分会场 卢顺杰 试验协同仿真- VIC-3D非接触全场应变测量系统.pdf",
            "sort": 2639,
            "attachid": "atwwevfuhbuw5s7y1tnegvev5tphaxkg",
            "type": "3"
        }, {
            "path": "doc/20190719231209856-轻量化技术分会场 曾锋 HyperWoks在小型客车轻量化设计中的应用.pdf",
            "intime": "2019-07-20 04:11:11",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "轻量化技术分会场 曾锋 HyperWoks在小型客车轻量化设计中的应用.pdf",
            "sort": 2638,
            "attachid": "atyl799f5lwxer5adtrs60byx8fz4ei6",
            "type": "3"
        }, {
            "path": "doc/20190719231208868-轻量化技术分会场 童忠财 新型尾门复合材料混合增强件优化设计.pdf",
            "intime": "2019-07-20 04:11:10",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "轻量化技术分会场 童忠财 新型尾门复合材料混合增强件优化设计.pdf",
            "sort": 2637,
            "attachid": "atmbn2gg5p47yozi1z3d7krvwntppycy",
            "type": "3"
        }, {
            "path": "doc/20190719231206862-轻量化技术分会场 冯玉龙 HyperWorks在无人机结构设计中的应用.pdf",
            "intime": "2019-07-20 04:11:08",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "轻量化技术分会场 冯玉龙 HyperWorks在无人机结构设计中的应用.pdf",
            "sort": 2636,
            "attachid": "ate9bzn2qk49up0igodaozg596hpsy8y",
            "type": "3"
        }, {
            "path": "doc/20190719230950861-主会场 James Dagg 基于HyperWorks的仿真驱动设计解决方案.pdf",
            "intime": "2019-07-20 04:08:52",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 James Dagg 基于HyperWorks的仿真驱动设计解决方案.pdf",
            "sort": 2635,
            "attachid": "at6f12ht4qiyg2he5l6c5y6cs3kpte4s",
            "type": "3"
        }, {
            "path": "doc/20190719230847865-主会场 周明 Altair优化技术驱动设计创新.pdf",
            "intime": "2019-07-20 04:07:49",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 周明 Altair优化技术驱动设计创新.pdf",
            "sort": 2634,
            "attachid": "atpu0o839wie6eetf2rgbhs8uyr9nqn3",
            "type": "3"
        }, {
            "path": "doc/20190719230841867-主会场 顾剑民 汽车行业三大变革与法雷奥创新技术.pdf",
            "intime": "2019-07-20 04:07:43",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 顾剑民 汽车行业三大变革与法雷奥创新技术.pdf",
            "sort": 2633,
            "attachid": "atht2kb8r1ywq24u7cy9ae8ufv71gu4u",
            "type": "3"
        }, {
            "path": "doc/20190719230834861-主会场 胡震东 数据为王-打造好用、易用、爱用的数据管理平台.pdf",
            "intime": "2019-07-20 04:07:36",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 胡震东 数据为王-打造好用、易用、爱用的数据管理平台.pdf",
            "sort": 2632,
            "attachid": "atl8taxzi4kbcwlcl3t595g6xyc7vafx",
            "type": "3"
        }, {
            "path": "doc/20190719230827868-主会场 宋廷伦 智能化数字化技术驱动汽车产品创新.pdf",
            "intime": "2019-07-20 04:07:29",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 宋廷伦 智能化数字化技术驱动汽车产品创新.pdf",
            "sort": 2631,
            "attachid": "atmykob7pv738z84535zacsdsdzoha03",
            "type": "3"
        }, {
            "path": "doc/20190719230820456-主会场 Doug Ellis Altair Knowledge Works Data Intelligence.pdf",
            "intime": "2019-07-20 04:07:21",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "主会场 Doug Ellis Altair Knowledge Works Data Intelligence.pdf",
            "sort": 2630,
            "attachid": "atvzvongvspggwag85hqc6aggsmzuoyi",
            "type": "3"
        }, {
            "path": "doc/20190719224246865-Academic Day 浦实 基于Altair电磁仿真的创新创业创造高校实践.pdf",
            "intime": "2019-07-20 03:41:48",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day 浦实 基于Altair电磁仿真的创新创业创造高校实践.pdf",
            "sort": 2629,
            "attachid": "atwb80h7is2hfmn3g8gsvw8vgkdx4upa",
            "type": "3"
        }, {
            "path": "doc/20190719224227854-Academic Day 周昊苏 OptiStruct在方程式赛车设计优化中的应用.pdf",
            "intime": "2019-07-20 03:41:29",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day 周昊苏 OptiStruct在方程式赛车设计优化中的应用.pdf",
            "sort": 2628,
            "attachid": "atgun3r7qg2gqp4tb8g779g403q048oi",
            "type": "3"
        }, {
            "path": "doc/20190719224223859-Academic Day 李翠超 汇聚solidThinking创新活力，打造增材制造精品课程——Altair与上海交大校企合作分享.pdf",
            "intime": "2019-07-20 03:41:25",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day 李翠超 汇聚solidThinking创新活力，打造增材制造精品课程——Altair与上海交大校企合作分享.pdf",
            "sort": 2626,
            "attachid": "at1qovligg243yqvw09lemdfqdn0p57o",
            "type": "3"
        }, {
            "path": "doc/20190719222709856-Academic Day 李征 Altair 高校赋能行动——合作战略与实践.pdf",
            "intime": "2019-07-20 03:26:11",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day 李征 Altair 高校赋能行动——合作战略与实践.pdf",
            "sort": 2625,
            "attachid": "at23cq25sdyo2mt8ubgbeazfp8uubg45",
            "type": "3"
        }, {
            "path": "doc/20190719222706857-Academic Day Carol Wu 台湾大学生创新设计案例赏析.pdf",
            "intime": "2019-07-20 03:26:08",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day Carol Wu 台湾大学生创新设计案例赏析.pdf",
            "sort": 2627,
            "attachid": "atnn3t23pbtbe5lcnbuff2ocsh12w0xu",
            "type": "3"
        }, {
            "path": "doc/20190719222700855-Academic Day Jordi Soler Altair EM Student Competition.pdf",
            "intime": "2019-07-20 03:26:02",
            "max": 0,
            "dtypeid": "dti2u56bnx0nqhibym3tsaeoo38z6vwo",
            "name": "Academic Day Jordi Soler Altair EM Student Competition.pdf",
            "sort": 2624,
            "attachid": "at09q0ps86w6kbhgw5zwoq5lsq3pmb66",
            "type": "3"
        }]
    }
}

items = json['data']['attachs']

url = "http://blog.altair.com.cn/file-handler/file-download/"

headers = {
    "Cookie": "JSESSIONID=CCE07B0785BE46B8710FCFC2F0CFF447"   ## 去浏览器里面抓
}

# 批量下载

for item in items:
    print(url + item['name'])
    file_url = url + "/" + item['attachid'] + "/" + item['name']
    r = requests.get(file_url,headers=headers)
    print(r.status_code)
    f = open(item['name'], "wb")
    for chunk in r.iter_content(chunk_size=512):
        if chunk:
            f.write(chunk)
