# GTM实习工作梳理

## 价格监控

价格监控有两个，一个是**每天**更新各网站小米产品的价格数据，一个是**每周五**更新竞品价格的数据

### 1\.1 小米产品每日价格监控

每天早上更新各网站小米产品价格数据在[promo tracking OM SMP \+ PAD](https://mi.feishu.cn/sheets/YpORspfhXhOUNStcGEbcghyunye?sheet=76SlJX)

复制前一天的sheet，修改名称成当天日期，对应更新价格

如网站价格和促销价一致则填写OK，如不一致则填写网站的实际价格。若网站没有库存或查不到这个品，那就写unavailable，若线上没有但是显示线下门店有，那就写unavailable online，总之就是根据实际情况填写。

如果不同颜色有不同价格，建议都写上，Eg\. 1699 black; 1629 brown; OK blue

（原文此处有截图，见飞书原文档）

每周一开始查价格之前，先根据[2026 OM promo offer](https://mi.feishu.cn/sheets/MjgvsQnv2hhF54tgxB6c4uBEnqe?sheet=hDZuV8)里面对应当周的promo price更新RRP after Promo列的价格并标颜色，对应修改offer start date和offer end date列日期

（原文此处有截图，见飞书原文档）

\*需要注意的是，我们不需要看outlet产品的价格！如果刷到outlet可以忽略！

此外，Allegro我们只看小米官方店的价格，留意一下是不是小米的店

（原文此处有截图，见飞书原文档）

以及MSH\(Media Market\)，只看MSH官方店的价格。下图为MSH官方店

（原文此处有截图，见飞书原文档）

下图为非官方店，看到sales and shipping by xxxx的字样，就不是MSH官方店

（原文此处有截图，见飞书原文档）



### 1\.2 每周五竞品价格监控

每周五更新在这个表[竞品价格 \- 2026](https://mi.feishu.cn/sheets/C9qSsv2mehberdt9jyjcPQLknwe)

在上周的数据左侧插入新的一列，填写对应产品在EURO和MEX的价格

如果两个网站价格一致则直接填写，如不一致则用/分隔开

（原文此处有截图，见飞书原文档）



### 1\.3 查价格VBA工具

可以使用下面Excel里的VBA工具，其中包含了目前价格监控的链接，如果后续有变动可以更新对应的产品和链接

\[XM Price Checker\- update\.xlsm\]

**小米产品的价格监控**，就用promo tracking sheet左上角的按钮，每次会打开两行包含的所有链接，然后再人工查看价格对比手动更新到promo tracking飞书表

（原文此处有截图，见飞书原文档）

查询完成关闭这两行链接之后，回到excel工具，点击continue，会继续打开下两行的链接，依此类推直到结束。

如果点击cancel就会取消这个进程，后面重新点按钮会从头开始

（原文此处有截图，见飞书原文档）

**竞品价格监控**就用竞品价格sheet，和上面的一个原理，但是每次会打开10行，以此类推。

（原文此处有截图，见飞书原文档）

\*如果需要修改这个按钮，可以先点击Developer，再右键点击按钮，选择Assign Macro\.\.

（原文此处有截图，见飞书原文档）

这里面有两个，一个是TwoRows每次打开两行，一个是TenRows每次打开10行，选择你想修改的点击编辑

（原文此处有截图，见飞书原文档）

就会进入到修改代码的地方，比如这个For r = 2, 就是这个打开链接的程序会从第二行开始运行，如果你中途不小心点了cancel想要继续但不想从头开始，可以修改这个行数，比如从第10行开始就改成r=10

（原文此处有截图，见飞书原文档）

修改好了之后保存关闭这个界面回到excel就可以正常使用修改后的按钮了



### 1\.4 查价格Python工具

代码GitHub链接：https://github\.com/GarlieLi/XM\_Price\_Checker\.git

需要先装好python和git，以及跑代码的软件\(例如VScode,Pycharm,etc\)

搜索cmd，打开powershell

（原文此处有截图，见飞书原文档）

确认一下python和git已经装好以及版本是什么（不一定和我一样，这个问题不大）

（原文此处有截图，见飞书原文档）

输入：cd xxxx（你想保存的路径，下图以Desktop为例）

接着输入：git clone https://github\.com/GarlieLi/XM\_Price\_Checker\.git

完成之后就会看到所有代码文件下载在你选择的路径处了

（原文此处有截图，见飞书原文档）

接着打开，以VScode为例，项目架构大概如图

（原文此处有截图，见飞书原文档）

以下以Pycharm为例。接下来我们需要设置一个虚拟环境，在terminal依次运行这两个代码
python \-m venv \.venv

\.venv\\Scripts\\Activate\.ps1

（原文此处有截图，见飞书原文档）

那么虚拟环境就设置好了，接下来升级一下pip，跑这个代码：python \-m pip install \-\-upgrade pip

等它下载，如果已经是最新的那就不需要下载了，完成之后我们下载需要用的package

pip install pandas openpyxl playwright jupyter ipykernel

（原文此处有截图，见飞书原文档）

装完之后装Playwright browser，用这个代码：python \-m playwright install chromium

（原文此处有截图，见飞书原文档）

等它装完就好了，跑一下这个代码，把我们的虚拟环境搞成kernel：

python \-m ipykernel install \-\-user \-\-name xm\-price\-checker \-\-display\-name "Python \(XM Price Checker\)"

右上角的kernel选择我们的虚拟环境 XM Price Checker\(\.venv\)

（原文此处有截图，见飞书原文档）

\*如果是VScode,还需要装两个extension：**Python** — Microsoft，**Jupyter** — Microsoft

那么现在就配置好了，在跑之前我们先打开文件夹里的excel

（原文此处有截图，见飞书原文档）

打开它，把Results sheet里面的促销价和offer start/end date更新，后续每周对应更新和飞书文档同步

（原文此处有截图，见飞书原文档）

然后可以大概看看，需要查的Links是不是都enable了，如果只查部分可以把不需要的改成no

（原文此处有截图，见飞书原文档）

保存**关掉**Excel后，可以开始准备跑了，需要运行的是名称为XM\_price\_checker的jupyter notebook，在开头的第一个cell里面把project folder的路径替换成你电脑实际这个文件夹的路径，然后点击Run all开始运行

（原文此处有截图，见飞书原文档）

稍等一会，它会不断有浏览器窗口弹出来打开然后关闭，那么就是开始跑了，放着等它跑完，大概需要10分钟多

（原文此处有截图，见飞书原文档）

跑完之后再打开excel，可以看到Results sheet里面现在已经写入了，大致检查一下，手动修正不对的地方

（原文此处有截图，见飞书原文档）

Rawdata sheet里面就是抓取的原始数据，可以用作检查参考

（原文此处有截图，见飞书原文档）

第二天再跑之前，可以把Raw data都删掉，不然它会累计往下写这个表就会越来越大，以及Results sheet代码写入价格的部分也先清空只保留最开始的框架



\*后续如果有产品要新增的话，对应加在Results, Products, Links sheet里面，Links要放该网站搜索某产品的界面的link，例如下图。不要放选择某个产品点进去之后的link\!\!\!

（原文此处有截图，见飞书原文档）

Ps,目前这个价格抓取没有全部覆盖，Euro和Oleole因为antibots暂时解决不了，GG和Micom是html和价格库存等的显示规则不一样我暂时跳过了，Allegro我没有尝试但是感觉这个可行，但是要注意区分小米官方店。

有兴趣的话可以继续开发，架构和原理可以看README文件



\*没有覆盖的网站我还是使用上面的VBA工具，python这个sheet只有python没有覆盖到的link，点击按钮使用

（原文此处有截图，见飞书原文档）



## PSI表维护

PSI表格[Poland PSI DB\+Channel](https://mi.feishu.cn/sheets/I6DSsZsyahx9aCtkjlecNi2ynad?sheet=BSJYhL) 每周更新，周三下班前把OP更新完，周四下午3点前发SO预实分析前，更新完所有KA数据，没收到的可以去问对应的负责同事，实在没收到也没关系后面再补。

Rolv/GG/Euro/Auchan 找 @Dominika Wichowska

MSH/Xkom 找 @Agata Kozak Piotrowska

MEX/KTR/MAX 找 @Maciej Świątek

Tmobile/Play 找 @Sheng Jun 盛俊

Plus/Orange 找 @Jet Wu 吴江涛

### 2\.1 DB ST Stock更新

每周一上午更新，数据在公邮pl\-psi@support\.mi\.com

一共有三个经销商: AB, ALSO, IM。ST就是Sell through,卖出的给KA数量，Stock就是库存，经销商自己的库存量。对应数据更新在PSI表的DB STOCK和DB ST sheet

（原文此处有截图，见飞书原文档）

首先**AB的ST**，打开表格，选择数据创建数据透视表

（原文此处有截图，见飞书原文档）

接着在数据透视表里，把llosc放到values，其他是rows\. 

在Design\-Report layout里选择Tabluar form和Repeat all item Lables

（原文此处有截图，见飞书原文档）

以及Subtotals里面选择Do not show subtotals

（原文此处有截图，见飞书原文档）

完成之后这个透视表大概就长这样，我们等会复制粘贴就可以了

（原文此处有截图，见飞书原文档）

在PSI表的STsheet里面，手写一下周数和经销商名称，接着把数据贴到对应的列，贴完之后，其他列都是公式，直接复制上周的公式拉下来就可以了。注意AB的reference和ALSO,IM不同，需要复制上周AB的对应公式。

（原文此处有截图，见飞书原文档）

**ALSO的ST**在sellout\_PL sheet，但是里面没有时间，用xlookup从IMEI\_Sellout拉数据，用EAN码和客户名匹配

（原文此处有截图，见飞书原文档）

**IM的ST**原始数据没有需要处理的，可以直接复制粘贴。

然后一样的，复制到PSI表ST对应的列就好。注意IM/ALSO和AB贴的reference不同列，公式也不同，贴完数据去复制上周IM/ALSO的公式直接拉就可以了。

（原文此处有截图，见飞书原文档）

接着是**STOCK**，这个三个经销商都差不多，我们直接找到对应表格复制对应的数据

AB的stock数据在这里找

（原文此处有截图，见飞书原文档）

ALSO的在表格Stock\_PL sheet

（原文此处有截图，见飞书原文档）

IM的数据在另一个邮件，如果公邮没有的话就去找@Davis Miao 缪鹏飞要

（原文此处有截图，见飞书原文档）

Stock的数据都差不多，也不需要处理原始数据，直接打开复制粘贴就可以，下图以AB为例

（原文此处有截图，见飞书原文档）

PSI表里面，手写Week和经销商名称，把excel数据贴到STOCK对应的列，剩下的公式拉齐

（原文此处有截图，见飞书原文档）



\*Stock和ST都是，有时候拉完公式可能会出现没有数据显示0的，这个情况先看看产品名是什么，如果是手环/手机壳/平板笔之类的产品那可以直接删掉。如果是新品以前没有的，那就需要更新引用表。

引用表是这个sheet，SKU等信息可以在[PL GTM All\-In\-One ](https://mi.feishu.cn/sheets/QA7LsH8PohM41ctYAcmcOGaSn7c?sheet=FwDsGH)这个表看

（原文此处有截图，见飞书原文档）

还是找不到的话可以问@Pope Peng

**更新完DB ST Stock在GTM群里说一声**



### 2\.2 KA周报更新

数据在公邮pl\-psi@support\.mi\.com，收到就更新，尽量在周四下午3点前完成

ALL,OM,OP这三个sheet一般不动，都是公式拉的。OM是所有公开市场，OP是四大运营商，ALL是OM\+OP数据。

其中 Michain 是 GG, ROLV 的总和；OM Others 是 MAX, Auchan, KTR 的总和

（原文此处有截图，见飞书原文档）

一般我们更新ST, SO, Channel Inventory Actual的数据，分别对应sell through, sell out, stock

（原文此处有截图，见飞书原文档）

#### 2\.2\.1 ST数据更新

ST用的是两个数据源，一个是公式从DB ST sheet里面拉，一个是在国际BI手动查询。

OM基本上公式拉就行，四大运营商在下面这个国际BI查，MEX需要拉公式之后在手动加上国际BI查的妥投量

国际BI的网址：https://iwork\.mioffice\.cn/idi/Main/Shujing/PSI?id=all\_sale\_dataset



先点击KPI\(sale\)Dataset, 把项目和内存拖到右侧维度，妥投量拖到右侧指标

（原文此处有截图，见飞书原文档）

将周拖动到筛选，选择静态区间，再选择对应需要查询的周，点击确定

（原文此处有截图，见飞书原文档）

再把客户名称拖到筛选，输入选择对应需要查询的公司主体名，下图以PLAY为例

（原文此处有截图，见飞书原文档）

点击妥投量从高到低排列，然后根据这个数据填写到对应PSI表上。

（原文此处有截图，见飞书原文档）

填写完之后记得标记颜色，妥投量0的不用标，**四大运营商的ST数据以国际BI为准**，如果表格上本来就有数据但是国际BI没查到的，那就改成0。如果当周没有ST，那就随便找个0标记一下颜色，表示ST已经更新了。

（原文此处有截图，见飞书原文档）

**OM**的ST是通过公式从DB ST拉的，先筛选ST，可以看到上周已经有sumifs公式了，直接整列复制粘贴到新的一周，接着再将上周的数据复制粘贴成数值，这样就完成了。

（原文此处有截图，见飞书原文档）

\*需要注意的是**MEX**，不仅公式从DBST拉数据，还需要在国际BI查

复制上周公式到本周，上周复制粘贴成数值，这两步和其他OM是一样的。复制完公式之后，需要检查公式里有没有在最后加数，例如下图，需要手动把这个加上去的数去掉。

（原文此处有截图，见飞书原文档）

确认完上周复制来的公式都干净之后，去国际BI查MEX的妥投量，按上文的方法。查到当周有数的，再手动加到公式后面，例如上图。



#### 2\.2\.2 OP周报更新

四大运营商有PLAY, PLUS, ORANGE, TMO，运营商的周报都是需要手填到PSI表

**Orange**可以找@Jet Wu 吴江涛要销售整理好的，如下图，就按照产品名在PSI表对应填，填完标颜色

（原文此处有截图，见飞书原文档）

**PLUS**同理，按产品名对应填到PSI表并标记颜色，都是0的不用填

（原文此处有截图，见飞书原文档）

在PSI表筛选SO和inventory actual一起填方便点，主要看上周有标记过颜色的跟着更新，以及留意刚上的新品

（原文此处有截图，见飞书原文档）

**PLAY**的周报用SKU匹配上周的表里的code，如果出现\#N/A，那就看看是什么东西，非手机类的可以不管，手机或者pad的话根据产品名手动写上code，有多个内存的产品注意区分内存

（原文此处有截图，见飞书原文档）

接着做数据透视表，再通过产品代码去PSI表手动对应更新

（原文此处有截图，见飞书原文档）

**T\-mobile**也是类似的，用material匹配上周周报里的code，有\#N/A的话自己判断手动补充

（原文此处有截图，见飞书原文档）

其中产品名称看到后面有Outlet字样的都是outlet产品，有DEMO的都是样板机，不需要统计这些的销量和库存，code就直接写Outlet或者Demos\. 如果遇到其他后缀不确定的可以问@Sheng Jun 盛俊

（原文此处有截图，见飞书原文档）

搞完code之后，创建数据透视表，再用产品代码去PSI表手动对应更新SO和stock

（原文此处有截图，见飞书原文档）



#### 2\.2\.3 OM周报更新

OM的周报更新原理都是一样的，大概都是用EAN码\(69码\)匹配产品代码，再按PSI表的顺序复制代码然后SUMIF，再复制粘贴回PSI表，具体详情如下。

首先需要用到的TTL工具表是这个，匹配的代码是K列code\+内存。有新品的时候更新这个表。

\[TTL ID 10\.7更新\.xlsx\]



以**EURO**为例，他们把Pad和手机分开两个邮件发的，我们需要把两个文件的数据汇总

（原文此处有截图，见飞书原文档）

（原文此处有截图，见飞书原文档）

在手机的文件里，还有手环和手表等，这些可以删掉

（原文此处有截图，见飞书原文档）

（原文此处有截图，见飞书原文档）

整理完就vlookup用EAN码去TTL工具表匹配code，检查是否有\#N/A，SUM求和总的库存和SO数

（原文此处有截图，见飞书原文档）

然后在PSI表EUROsheet里面筛选SO，从Others下面开始复制整列产品代码，贴到刚刚的Excel里

（原文此处有截图，见飞书原文档）

接着SUMIF，用代码匹配，分别求和SO和Stock

（原文此处有截图，见飞书原文档）

公式拉完之后同理，最下方再用SUM求和一下所有的SO和stock

然后把刚刚原始数据SUM的库存和销量数据写过来，相减，得出的数就是Others

（原文此处有截图，见飞书原文档）

接着就回到PSI表，分别筛选SO和inventory actual，手填others的数据，下面整列复制excel里刚刚SUMIF的结果，粘贴为数值，标记颜色和字体，这样就完成了。粘贴完可以快速复核一下，整列全选看sum的总数是否一致。

（原文此处有截图，见飞书原文档）



**MEX**的邮箱带TERG，他们的表数据很多，先筛选SMARTFONY和TABLETY，复制到一个新的sheet再开始做

（原文此处有截图，见飞书原文档）

后续步骤和上面一样，做完SUMIF复制粘贴整列到PSI，手填others

（原文此处有截图，见飞书原文档）



**MSH**的邮件后缀是media\-saturn，数据看Total sheet,也是有很多其他品类的，都删掉只保留手机和平板

（原文此处有截图，见飞书原文档）

后续步骤都一样，做完SUMIF整列复制粘贴到PSI，手填others

（原文此处有截图，见飞书原文档）



**Xkom**也是一样，就是需要留意他们把所有week都放在表格里，所以SUMIF的区域注意只看要做的当周

（原文此处有截图，见飞书原文档）



**OM others **由 Auchan,KTR\(邮箱后缀komputronik\),MAX\(邮箱后缀gtpoland\)组成，先分别用TTL工具表匹配代码

（原文此处有截图，见飞书原文档）

注意KTR，有些EAN码没有开头没有6，写个公式补上，粘贴成数值，再用修复好的69码去TTL工具表匹配code

（原文此处有截图，见飞书原文档）

以及筛选一下品类，只选手机和平板

（原文此处有截图，见飞书原文档）



这几个分别匹配完没问题就筛选除了\#N/A的，粘贴汇总到一个新的excel，再按之前的操作SUMIF再贴PSI

（原文此处有截图，见飞书原文档）



**Michain**是GG和ROLV的汇总，也是先分别在表格里找productID用TTL工具表匹配代码，以及在图中这里筛选T

（原文此处有截图，见飞书原文档）

然后汇总，因为GG和ROLV都把SO和stock分开两个表，所以汇总的时候分别贴，接着一样的SUMIF再贴PSI

（原文此处有截图，见飞书原文档）



### 2\.3 PSI表增加产品

有新品的时候，根据GTM的需求加到PSI表上，例如某产品加到所有OM某产品加所有OP，根据实际情况加

注意**所有sheet的产品代码保持一致**，比如Redmi Note 17 Pro 5G是Q16也是MOSCOW，都按系统写MOSCOW



以EURO为例，每个产品是有8行，可以直接选择插入8行，复制上一个产品的I列粘贴，其他根据实际信息填写

（原文此处有截图，见飞书原文档）

需要留意**F列是公式**，按之前的公式，不要手写这个产品代码。以及注意G列H列数据前后**不要有空格**，不然F列汇总的代码也会有空格，会影响到公式拉取数据。

（原文此处有截图，见飞书原文档）

以及可以看到后面的月份都有预设好的公式，直接复制一个之前产品的所有公式粘贴就行

**注意，除了各KA，也需要同步加到OM/OP和ALL的sheet\!**



## SO预实分析

每周四下午3点前更新[波兰公开市场 SO预实分析](https://mi.feishu.cn/sheets/H6TAsLGLehw6jEtdZtDcfzsUnZd?sheet=Em3fP4)

复制上周的sheet修改名称为本周

（原文此处有截图，见飞书原文档）

在对比列的左侧，插入一列写本周，修改单元格名称，公式直接复制粘贴上周的

（原文此处有截图，见飞书原文档）

修改对比列的名称和公式，变成本周/上周\-1，就做完了

（原文此处有截图，见飞书原文档）

这个表的数据是拉PSI表的，所以如果有KA的周报还没更新到PSI，这个表的数据就会不完整。尽量把所有KA数据都在周四下午3点前更新完到PSI，实在没有的话发SO预实分析的时候说一下缺哪些数据。



**更新完发在GTM群里**，大概周四下午2\-3点发，说明缺了什么数据（如有）

（原文此处有截图，见飞书原文档）



## 联合营销/会议费相关

总览可以参考这个文档 [Joint Marketing projects approval process](https://mi.feishu.cn/wiki/H9mhwkARJiEjQrkxziWcIsv1n2f)

联合营销和会议费流程具体操作看这个 [波兰\-联合营销/会议费提单SOP](https://mi.feishu.cn/wiki/LGjzwhVrKiOM6OkPC2Ac70hxn3d)

联合营销验收现在不走CRM系统，除了Q2还没结的单子会有销售找你，其他的均按照下面的SOP进行：

（原文此处有截图，见飞书原文档）

（原文此处有截图，见飞书原文档）

- 供应商名称：

    - ORANGE: ORANGE Polska S\.A\.      CRM20231205001

    - PLAY: P4 Sp\. zo\.o\.        CRM20231130001

    - T\-Mobile: T\-Mobile Polska S\.A\.    CRM20240603001

    - PLUS : Polkomtel sp\. zo\. o\.   CRM20220523001

    MEX: INGRAM MICRO SP\. ZO\. O\.

    Euro: INGRAM MICRO SP\. ZO\. O\.

    MSH: Also Polska Sp\. zo\. o\.

    Neonet: INGRAM MICRO SP\. ZO\. O\.

    GGmedia/Rolv: INGRAM MICRO SP\. ZO\. O\.

    

TMO：T\-mobile Polska, 二级客户不填, 100%，

PLAY：P4 Sp\.ZO\.O\(检查全称）,二级客户不填，100%

PLUS: Polkmotel SP\.ZO\.O, 二级客户不填，100%；

Orange：Orange S\.A, 二级客户不填，100%；

MEX：Ingram，二级客户Terg，100%；

Euro：Ingram，二级客户Euro，100%；

MSH：Also，二级客户Media Markt，100%；

Neonet：Ingram，二级客户Neonet，100%；

GGmedia/Rolv：Ingram，二级客户分别选GGmedia或rolv，100%；







发票录入和对公付款在这里 [对公付款提单参考\-联合营销流程](https://mi.feishu.cn/docx/WpLZdFrfioRRXCxplT7cusnJngf)

Q3之后的联合营销都是合同返利，在小米审批走操盘费申请，详情见下图

CN : 联合营销CN支付，走合同返利池。

（原文此处有截图，见飞书原文档）

供应商名称：

对方名称：\(包阿茹罕\(Mia\) 这里检查各公司全称并补充，需仔细检查空格和点\)

- ORANGE: ORANGE Polska S\.A\.      CRM20231205001

- PLAY: P4 Sp\. zo\.o\.        CRM20231130001

- T\-Mobile: T\-Mobile Polska S\.A\.    CRM20240603001

- PLUS : Polkomtel sp\. zo\. o\.   CRM20220523001

MEX: INGRAM MICRO SP\. ZO\. O\.

Euro: INGRAM MICRO SP\. ZO\. O\.

MSH: Also Polska Sp\. zo\. o\.

Neonet: INGRAM MICRO SP\. ZO\. O\.

GGmedia/Rolv: INGRAM MICRO SP\. ZO\. O\.



供应商编码查询方法：[客户 \- 国际统一工作台](https://iwork.mioffice.cn/main-global-customer-mgmt/customerManage/admissionChange)

查询网址：https://iwork\.mioffice\.cn/main\-global\-customer\-mgmt/customerManage/admissionChange

（原文此处有截图，见飞书原文档）



提交前在备注写一下项目名称，SO达成情况等。





## Offer录入和发放

具体流程看[波兰渠道财经公开市场\&Others Offer结算SOP](https://mi.feishu.cn/wiki/PifSwtpMXii0NKk2FxecPG0in9x) 

二、预算申请及offer发放 里面的 3 \.系统政策录入 以及 4 \.系统政策发放



在系统录入offer的时候，注意这几列**不要复制粘贴**，这些是填了其他信息系统自己带出来的。

（原文此处有截图，见飞书原文档）



\*留意一下**P83**这个产品，填完项目代码和ram带出来的产品名称经常是with keyboard, 录完要把这一列拉开看一下名称，如果是with keyboard就手动修改，把with keyboard删掉，只保留Redmi Pad 2 Pro

（原文此处有截图，见飞书原文档）



## OM offer的预测销量拆分

**每个月月底**，等财务通知，需要做OM offer的预测销量拆分，填写到BOC表的L列[BOC offer budget](https://mi.feishu.cn/wiki/EcoWwZXDAi6xVrknZBbcRKEcn8e?sheet=d7b511)

（原文此处有截图，见飞书原文档）

预测销量的数据来源于这个表[2026 OM promo offer](https://mi.feishu.cn/sheets/MjgvsQnv2hhF54tgxB6c4uBEnqe?sheet=4GMnli)

每个产品SO\-GTM就是对应每周的预测销量

（原文此处有截图，见飞书原文档）

往后拉还有客户拆分比例，我们就是要把周的预测销量按比例拆分到具体的客户

（原文此处有截图，见飞书原文档）



可以创建一个在线表格，让虾来做拆分，以这个为例[Oct offer 拆分](https://mi.feishu.cn/wiki/Rw19wgm2pikTjwk7UbRcBJhun6e?sheet=45d4f2)

把BOC的offer信息放到sheet1

（原文此处有截图，见飞书原文档）

在OM promo offer表提取SO\-GTM信息和对应的客户拆分比例，放到sheet2

（原文此处有截图，见飞书原文档）

然后可以用这个prompt发给虾⬇



帮我完成这个offer拆分，总共有300多行需要完成，具体如下：

请帮我处理我上传的 Excel 文件，完成 Sheet 1 的 L 列预测销量拆分。

**一、两个 Sheet 的字段含义**

**Sheet 1**

需要填写的是 **L 列**。

关键字段：

- C列 = 客户名字

- I列 = 产品代码

- J列 = 内存

- O列 = 预测开始日期/期间开始日期

- P列 = 预测结束日期/期间结束日期

其中 O列和P列定义了这个客户/产品在什么时间范围内需要计算预测销量。

**Sheet 2**

Sheet 2 是每周的总预测销量以及客户占比。

关键字段：

- B列 = 产品代码

- C列 = 内存

- E1:I1 = 周期，分别对应 W36、W37、W38、W39、W40 的具体日期

- E列:I列 = 每个产品\+内存组合在对应周的总预测销量

- M:U列 = 客户占比

Sheet 2 中的客户名称需要根据表头确定，M:U 每一列分别对应一个客户。

**二、具体计算逻辑**

请严格按照下面的逻辑计算 Sheet 1 的 L 列，不要直接简单做 SUMIFS。

**第一步：先计算 Sheet 2 每个客户每周的预测销量**

对于 Sheet 2 的每一行：

首先根据：

- 产品代码 = Sheet 2 B列

- 内存 = Sheet 2 C列

确定这个产品\+内存组合。

例如某一行：

- 产品代码 = ABC123

- 内存 = 256GB

- E3 = 650

那么 E3 的 650 就代表这个产品\+内存在 W36 这一周的总预测销量。

然后根据 M:U 中对应客户的比例进行拆分。

例如：

- E3 = 650

- 客户 A 在 M列的比例 = 20%

那么：

**客户A W36预测销量 = 650 × 20% = 130**

同理：

**客户某周预测销量 = 该产品\+内存该周总销量 × 该客户在 M:U 对应列的比例**

需要对：

- W36

- W37

- W38

- W39

- W40

分别计算。

最终得到类似这样的逻辑：

产品代码

内存

客户

W36

W37

W38

W39

W40

ABC123

256GB

客户A

130

150

120

140

160

**三、第二步：根据 Sheet 1 O列和P列确定需要哪些周**

然后处理 Sheet 1。

对于 Sheet 1 每一行，根据：

- C列 = 客户

- I列 = 产品代码

- J列 = 内存

- O列 = 开始日期

- P列 = 结束日期

判断这个日期区间覆盖 W36\-W40 中的哪些周。

这里要注意：

**Sheet 2 的 E1:I1 是每个周对应的具体日期，所以请使用 Sheet 2 E1:I1 的实际日期来判断，而不是把 W36、W37 等当成普通文本。**

例如：

如果某个 Sheet 1 行：

- 客户 = 客户A

- 产品代码 = ABC123

- 内存 = 256GB

- O列开始日期 = W36 对应日期

- P列结束日期 = W38 对应日期

那么应该计算：

**客户A W36 \+ 客户A W37 \+ 客户A W38**

假设：

- W36 = 130

- W37 = 150

- W38 = 120

那么 Sheet 1 L列最终结果：

**L = 130 \+ 150 \+ 120 = 400**

**四、第三步：匹配条件**

计算 Sheet 1 L列时，必须同时满足以下三个条件：

1. Sheet 1 I列产品代码 = Sheet 2 B列产品代码

2. Sheet 1 J列内存 = Sheet 2 C列内存

3. Sheet 1 C列客户名字 = Sheet 2 M:U 中对应的客户名称

只有这三个条件全部匹配，才能取得正确的客户周预测销量。

然后再根据 Sheet 1 O列和P列的日期范围，筛选出需要累加的周。

**五、最终公式逻辑**

对于 Sheet 1 每一行：

**L列 = Σ（Sheet 2 对应产品\+内存在目标周的总销量 × 对应客户比例）**

其中 Σ 的范围由 Sheet 1 O列和P列决定。

也就是说：

**L = 所有满足 O列 ≤ 周日期 ≤ P列 的周的客户预测销量之和**

**六、非常重要的细节**

请你先检查 Sheet 2 的结构，确认：

- E1:I1 确实是 W36\-W40 对应的日期

- M:U 的表头确实是客户名称

- E:I 是 W36\-W40 的总销量

- B列和C列分别是产品代码和内存

如果实际 Excel 中表头所在行、客户名称所在行、日期格式等和上述描述有一点不同，请以实际 Excel 内容为准，不要擅自假设。

另外：

- 产品代码可能有文本格式，不能因为看起来像数字就转换成数字。

- 内存也要按照原始值精确匹配，例如 128GB、256GB 等。

- 客户名称要精确匹配，注意前后空格。

- O/P 日期如果包含时间，也要正确处理。

- 如果某个产品\+内存\+客户没有对应比例，则该周销量按 0 处理。

- 如果比例为空，也按 0 处理。

- 如果某周不在 O列:P列范围内，则不要计入。

- 如果 O列/P列为空，请不要随意计算，先标记出来。

- 如果找不到匹配的产品代码\+内存，也请标记出来，而不是默默填 0。

- 不要修改 Sheet 1、Sheet 2 的原始数据和格式，只填写 Sheet 1 的 L列。

**七、请实际执行**

请直接打开并分析 Excel 文件。

完成以下步骤：

1. 读取 Sheet 1 和 Sheet 2 的实际结构。

2. 自动识别 Sheet 2 E1:I1 对应的 W36\-W40 日期。

3. 建立“产品代码 \+ 内存 \+ 客户 \+ 周”的预测销量。

4. 根据 Sheet 1 每一行的 C、I、J、O、P 找到对应的客户、产品、内存和日期范围。

5. 将对应周的客户预测销量加总。

6. 将最终结果写入 Sheet 1 的 L列。

7. 保持原有 Excel 格式不变。

8. 完成后检查 L列结果，特别检查几个样例行，确认计算逻辑正确。

9. 最后告诉我： 

- 一共处理了多少行；

- 成功填写多少行；

- 有多少行因为产品/内存/客户/日期无法匹配而需要人工检查；

- 随机列出 3\-5 行计算过程，让我可以核对结果。

请不要只告诉我应该怎么做，**直接对 Excel 文件执行计算并填写 Sheet 1 的 L列。**



⬆发给虾的时候，记得要附上offer拆分的飞书文档，还要给虾编辑权限

（原文此处有截图，见飞书原文档）

虾拆分完之后再去大概检查一下，OK的话就填到BOC表然后喊@Evelyn Zhang 张小艺看拆分结果



## 总览

数据类：

每周一上午更新DB ST/Stock

每天上午更新价格监控

每周四下午三点前完成更新PSI和SO预实分析

每周五竞品价格查询



会议费：

Meeting fee表有Alex approval之后走预算申请，

预算通过之后根据模板建合同，在合同管理系统申请，

合同审批通过后打印拿给迪总签字（找@HuangLin转交，琳姐不在找@Konstan 黄聪），签完拿给对应销售

双签完成后填到这里[海外合同过用印节点申请表Overseas Contract Approval Form for Stamp Application](https://mi.feishu.cn/wiki/XIrPwQtRPidowskmWV4cfOiinrG?table=tblNWesI9ngOqNjq&view=vewoP9kTVk)，

活动结束有proof后，从预算里面点去付款走对公付款



联合营销：

活动approve后邮件发offer，附上mediaplan发公邮和销售，确认后用公邮发客户

每月20号之前，根据[CEE 2026 Co\-mkt projects \- weekly review \- PL \- 区域GTM审批用](https://mi.feishu.cn/sheets/Fj7JshSXXh9iQstys5cc4mV0nUg?sheet=Lyq70V)的数据填[波兰联合营销冲减收入台账](https://mi.feishu.cn/sheets/G30AsmNSDh3CWhtRIArcMzWzn4g)，记得offer邮件/mediaplan/CEE表审批通过截图打包传attachment

活动结束有proof后走操盘费申请提验收，

如果是发票的，操盘费通过后发起对公付款



offer录入和发放\(激励系统\)：

月底做预测销量拆分，发OM的offer，留意有些客户pad和phone收件人不一样

其他有人找就对应录入系统，找人审批和发客户，记得同步更新BOC表的单号状态等

财务有时候会找修改预算，撤回对应的offer，修改预测销量，改到总投入达到财务的需求，找阳姐和小艺姐审批完后，创建offer但不用发给客户



