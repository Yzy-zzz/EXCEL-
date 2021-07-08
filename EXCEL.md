# EXCEL

### excel入门

#### 第一节

1.双击可以调整宽度

![image-20210704151953130](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704151953130.png)

2.合并后居中

![image-20210704152233344](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704152233344.png)

3.选中两列进行宽度高度调整就可以一起改变

4.序列填充

5.ctrl + ;	当前日期

6.从下拉列表选择：alt + 【下方向键】

#### 表格设计&自动加函数

1.列印，打印的时候默认周围是黑色，中间是虚线

![image-20210704153206389](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704153206389.png)

2.边框的绘制 按住ctrl 内部同样有框线

3.插入背景、设置网格线可见与否![image-20210704153718354](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704153718354.png)

#### 冻结窗格和分割视窗

1.冻结的是选中行的上一层

![image-20210704154251510](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704154251510.png)

2.拆分——分割视窗

#### 资料排序

比较简单

可以学一学自定义排序里面的东西

![image-20210704154857518](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704154857518.png)

#### 资料筛选

把符合条件的东西显示出来

高级筛选里面有很多不同的条件	可以自己去试试

#### 格式化为表格&交叉分析筛选器

#### 设置格式化表格的条件

![image-20210704164845880](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704164845880.png)

有 图阶、指示集、指示条等等

#### 工作表设定&合并计算

![image-20210704165308166](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704165308166.png)

#### 图表制作

![image-20210704165617134](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704165617134.png)

多玩

高级用法

![image-20210704170038147](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704170038147.png)







![image-20210704170044323](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704170044323.png	)             			               

组合图 ...

#### 数据透视图

可以看一下不同的行列情况下的情况

#### 打印问题

分页预览

![image-20210704182737095](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704182737095.png)

还能插入分页符

![image-20210704182907171](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704182907171.png)

设置打印的一些情况

![image-20210704183139949](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704183139949.png)

#### 页首&页尾&水印

![image-20210704183556125](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704183556125.png)

页首 页尾的格式设计

水印也是在页首 页尾插入到图片

### 函数

##### if函数

![image-20210704192742897](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704192742897.png)

巢状 IF

=IF(C7>60,"是","否")

=IF(C7>=90,"A",IF(C7>=80,"B","C"))

##### VLOOKUP

最左栏是主键【对于表中的结构要求】

Lookup_value是检索值

Table_array是检索范围

Col_index_num是第几列的数据

Range_lookup是 是否模糊查询【True or False】



![image-20210704193802643](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704193802643.png)

![image-20210704193728731](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704193728731.png)

##### 绝对参照

F4，将范围锁定

#### IFERROR

=IFERROR(VLOOKUP($C$3,$E$2:$K$12,7,FALSE),"查无此人")

#### 数据验证

![image-20210704203143305](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704203143305.png)

#### COUNTIF

两个参数 第一个范围 第二个【条件】

=COUNTIF(D2:D14,"<200")

=COUNTIF(D2:D14,"<"&F2)

加强版

【范围】【条件】【范围】【条件】

=COUNTIF(D2:D14,"<"&F2)

#### SUNIFS

条件加和

#### 名称管理

![image-20210704204614245](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704204614245.png)

可以给数据范围加名字

#### INDIRECT

根据自己之前定义的名字来提供可供选择的数据

### 一些操作技巧

1.分列

![image-20210704205750598](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704205750598.png)

2.ctrl + 【上下左右】可以直接跳到excel里面的四个角

​	在这个基础上+shift 可以进行选取

3.插入多个空白列

​	【F4】重复上一条指令

4.选中一行

ctrl 拖动是复制

shift + ctrl 拖动是 粘贴插入进去

5.移除重复项

![image-20210704210559310](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704210559310.png)

6.选择性贴上

里面可以转置

一起加或者一起剪

7.显示公式按钮

8.alt + enter 可以在同一栏回车

![image-20210704211159351](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704211159351.png)

### 数值格式

### 日期

### 成绩比较

#### RANK.EQ 函数

第三个输入方式，决定升序还是降序	1 0

#### RANK.AVG 函数

### 提取资料

#### LEFT

#### RIGHT

#### MID

#### FIND

FIND("搜寻文字"，“资料来源”)

FIND("搜寻文字"，“资料来源”，“起始位置”)

以上几个有机结合，就很牛

### HLOOKUP

按照行来

![image-20210704220942369](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704220942369.png)

### INDEX&MATCH

INDEX：回传数据

INDEX（范围，行数，列数）

MATCH:找到位置

MATCH("信息"，范围，匹配方式)

1 小于 0 等于 -1 完全小于

### 保护数据

![image-20210704222339416](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704222339416.png)

可以对一些行 列 工作sheet 进行隐藏	然后设置密码

也能对一些公式进行 隐藏

![image-20210704222639170](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704222639170.png)

还能对一些 特定区域的数据 进行编辑

### 重复资料

指相同的元组

那么可以进行 标注 或 删除

防止的时候 也用 资料验证

### 随机数

#### RANDBETWEEN(最小值，最大值)

按F9 重复运算

可以搭配 INDEX 使用

![image-20210704224337861](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704224337861.png)

CHOOSE 省去 建立辅助表格的麻烦

CHOOSE(index,"c1","c2","c3")。。。。

#### RAND

先rand12个 然后看每一个数在其中的排名，

然后除以12 的一半 分成了两部分，一个大于1 一个小于等于1

接着用向上取整 ROUNDUP 函数

最后搭配 CHOOSE

**然后把他们确定下来 要不每次操作都会发生变化**

### 做计划表

[【全套】Excel零基础入门进阶到函数，Excel自学教程从小白到高手超详细实操教程（Excel教程、Excel小白入门起步、Excel函数、Excel技巧）_哔哩哔哩_bilibili](https://www.bilibili.com/video/BV184411C7Ci?p=29&spm_id_from=pageDriver)

p29

### 甘特图

![image-20210704225748976](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704225748976.png)

### 柏拉图

![image-20210704231348249](C:\Users\我真好看\AppData\Roaming\Typora\typora-user-images\image-20210704231348249.png)

2/8	法则

2016版本之后 直接插入

### 宏 编程

我们宏可真是太厉害了

















