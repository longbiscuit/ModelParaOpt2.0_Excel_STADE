# ModelParaOpt2.0_Excel_STADE
EB, MC, MCC and CSUH model parameters were determined by non-inversion or inversion methods in Excel software.

relevant papers：

[1]Zhu B, Chen Z. Calibrating and validating a soil constitutive model through conventional triaxial tests: an in-depth study on CSUH model[J]. Acta Geotechnica, 2022: 1-14.

[2]Binglong Z, Xiaoxue S, Yuqing C, et al. Determining parameters of the CSUH constitutive model by genetic algorithm[J]. Japanese Geotechnical Society Special Publication, 2020, 8(6): 188-193.
# 简介
针对把耦合优化算法搭建在Excel平台上时存在反演速度过慢的问题，提出了分段调整模型迭代增量步长的动态迭代方法，实现了模型参数在广泛使用的通用软件Excel上的快速反演计算，对本构模型参数的优化确定提供了极大的便利。Excel通用软件平台操作简便且应用广泛，但其内嵌的编程语言计算速度较慢，加之反演参数过程需要大量的本构模型数值计算增量式的迭代计算，耗时太长。动态迭代方法将整个轴向应变按应力应变曲线斜率分为两段，曲线斜率较大段的轴向应变增量在迭代时呈步长较小的等差数列分布，而曲线斜率较小段的轴向应变增量在迭代时呈步长较大的均匀数列分布。动态迭代方法在保证预测精度的前提下极大的减少了模型迭代次数，使本构模型参数的反演在基于Excel的优化反演分析软件上实现成为可能。模型参数优化反演分析软件包含了粘土和砂土的统一硬化模型（CSUH, Unified hardening model for clays and sands）、修正剑桥模型（MCC, Modified Cam-clay model）、莫尔-库伦模型（MC, Mohr-Coulomb model）和邓肯-张模型（EB, Duncan-Chang model）几种典型本构模型参数的非反演确定方法和反演确定方法，非反演方法确定的模型参数也可以作为反演方法确定参数时的初始值以提高反演精度和速度。


# 页面展示
![动态迭代方法](https://user-images.githubusercontent.com/21994802/187571489-902b125a-bfc8-4388-a81c-cfac1765ae4b.jpg)

均匀迭代增量步长和等差数列分布的迭代增量步长计算结果对比


![功能](https://user-images.githubusercontent.com/21994802/187573678-5f8c55a8-0b7d-49cc-807e-9e3eb8e3fb8a.png)

参数优化反演分析软件（ModelParaOpt）的主要功能



![首页3](https://user-images.githubusercontent.com/21994802/188268917-94a1389a-c398-4537-ad81-b53b6da95e34.png)
参数优化反演分析软件（ModelParaOpt）首页


![过程页](https://user-images.githubusercontent.com/21994802/187611586-00411bef-892d-4476-adaa-d2ab73316657.png)

参数优化反演分析软件（ModelParaOpt）优化过程页




![敏感性分析](https://user-images.githubusercontent.com/21994802/187572219-bf92a28c-c22f-46b1-abb5-b1145c5bd79a.png)

参数优化反演分析软件（ModelParaOpt）敏感性分析页





![ISO](https://user-images.githubusercontent.com/21994802/187572265-5445e287-d44a-4b9a-8f9c-41e90f2870cf.png)

参数优化反演分析软件（ModelParaOpt）自动绘制对比图页（等向压缩试验）




![CD](https://user-images.githubusercontent.com/21994802/187572373-b8007089-8071-4d75-a80b-da910d3a1ff3.png)

参数优化反演分析软件（ModelParaOpt）自动绘制对比图页（CD试验）




![CU试验对比图](https://user-images.githubusercontent.com/21994802/187572419-d870e147-0ebc-4f69-9be7-f7f257a8a477.png)

参数优化反演分析软件（ModelParaOpt）自动绘制对比图页（CU试验）


![28a4c8487c28059090163ce3f3dad5a](https://user-images.githubusercontent.com/21994802/187614889-3b67ce2c-0fb0-4ad6-a319-3a0ff6fe256e.jpg)
代码模块（更具体的请看Excel的vba代码）

# 环境
## 运行环境:
'系统环境:
'Microsoft Windows 10 专业版(64 位) Build 19042 v10.0.19042
'Intel(R) Core(TM) i5-8250U CPU @ 1.60GHz, 1.80GHz, Dell Inc. 0XW4HD
'Intel(R) UHD Graphics 620, 1024MB VRAM, Driver v27.20.100.9316
'DDR4-2400 SDRAM:6.2/7.8GB
'C:188.4/249GB

## 软件环境:
'1-Excel 2019(开发环境)

'2-WPS教育版 -正式版 v11.3.0.8858-release(运行需安装vba_for_wps)
'C:\Users\Zhu Binglong\AppData\Local\Kingsoft\WPS Office\11.3.0.8858

在Excel和WPS中都可以运行。

# 注意事项
## 在C盘或桌面无法运行的问题
注意，如果将此excel放在桌面或C盘任何地方执行，可能由于没有C盘读写权限而发生错误：
![image](https://user-images.githubusercontent.com/21994802/197440367-b5a384d0-a939-4284-83ef-a5715bf96836.png)
![image](https://user-images.githubusercontent.com/21994802/197440461-cbf619dc-8928-4ba1-a98e-273310aac4bf.png)
解决方法有3种：
1.建议将此excel放在C盘以外的盘内运行。

2.如果只有C盘，可以以管理员权限打开此excel运行。

3.参考[解决vba无写C盘权限问题](https://blog.csdn.net/bobyeoh/article/details/9008005?spm=1001.2101.3001.6650.14&utm_medium=distribute.pc_relevant.none-task-blog-2%7Edefault%7EBlogCommendFromBaidu%7ERate-14-9008005-blog-88647398.pc_relevant_recovery_v2&depth_1-utm_source=distribute.pc_relevant.none-task-blog-2%7Edefault%7EBlogCommendFromBaidu%7ERate-14-9008005-blog-88647398.pc_relevant_recovery_v2&utm_relevant_index=14)
即用管理员运行cmd 输入`icacls c:\ /setintegritylevel M `解决。

