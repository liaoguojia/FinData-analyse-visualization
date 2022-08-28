import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar      #柱状图
from pyecharts.charts import Line     #折线图
from pyecharts.charts import Funnel   #漏斗图
from pyecharts.charts import Page
from pyecharts.globals import ThemeType
from pyecharts.charts import Grid
from pyecharts.commons.utils import JsCode 

## 数据导入与清洗
finRatio = pd.read_excel('./所有公司财务比率.xls')
finIndex = pd.read_excel('./所有公司财务指标.xls')

# 修改列名
finRatio.rename(columns={'上市公司代码_Comcd': 'comCd','最新公司全称_Lcomnm':'comNm','截止日期_Enddt':'endDt'},inplace=True)
finIndex.rename(columns={'上市公司代码_Comcd':'comCd','最新公司全称_Lcomnm':'comNm','截止日期_EndDt':'endDt'},inplace=True)
# 将两个表合成一个
finData =  pd.merge(finRatio,finIndex, how='inner', on=['comCd','comNm','endDt'])
finData

# 数据清洗
# 去除空值的行
finData.dropna(how='any',axis=0,inplace=True) 

for i in finData.comCd:
    if finData.comCd.to_list().count(i) != 5:    #去除数据不满五年的公司
        finData.drop(finData[finData.comCd == i].index, axis=0,inplace=True)

comNumber = int(len(finData)/5)  #行业公司总数
year_lst = list(finData.endDt.apply(lambda x:str(x)[:4]).unique())      #选取的五个年份

code = "C300750"       #本公司代码，可通过RPA接口传参 
name = "宁德"         #本公司名，可通过RPA接口传参 

finData.reset_index(inplace=True)  #设置新索引
finData.drop("index",axis=1,inplace=True)


## 本公司基本面分析
# 1-L 费用率表
expense = (
        Bar(init_opts=opts.InitOpts(width="700px", height="400px",theme= ThemeType.LIGHT ))#设置图表画布宽度 
        .add_xaxis(year_lst)
        .add_yaxis("销售费用率", finData[finData.comCd == code]["销售费用率(%)_Opeexprt"].to_list() ,stack="stack",bar_min_width=10,bar_max_width=50)
        .add_yaxis("管理费用率", finData[finData.comCd == code]["管理费用率(%)_Admexprt"].to_list() ,stack="stack")
        .add_yaxis("财务费用率", finData[finData.comCd == code]["财务费用率(%)_Finexprt"].to_list() ,stack="stack")
        #设置标签属性
        .set_series_opts(
            label_opts=opts.LabelOpts(position="inside",  formatter="{c}%")
            )     
        .set_global_opts(
            title_opts=opts.TitleOpts(title=name+"费用率表"),
            legend_opts=opts.LegendOpts(textstyle_opts=opts.LabelOpts(font_weight='bold')),#设置图例属性
            #设置横纵坐标属性
            xaxis_opts=opts.AxisOpts(name="Year",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold')),  #,interval=115,boundary_gap=['50%', '80%']
            yaxis_opts=opts.AxisOpts(name="Percent",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold'),axislabel_opts=opts.LabelOpts(formatter="{value}%"))
            )
)
expense.render_notebook()

def drawLineChart(df,cpnCode,chartName,x,y):
    """
    绘制折线图：x横坐标(年份)，y为展示的几项指标
    cpnCode：本公司代码
    chartName：图表名称
    """
    lineChart = Line(init_opts=opts.InitOpts(width="700px", height="400px",))   # 初始化折线图，设置宽高
    lineChart.add_xaxis(x)    #添加x轴 ：[str]

    for idx in y:   #添加所有y轴系列
        idx_ch = ""
        for i in idx:
            if("\u4e00"<i<"\u9fa5"):
                idx_ch +=i
        lineChart.add_yaxis(idx_ch,df[df.comCd == cpnCode][idx].to_list())  
        
    lineChart.set_global_opts(   #设置图表全局属性
        title_opts=opts.TitleOpts(title=chartName),   #图表标题
        legend_opts=opts.LegendOpts(textstyle_opts=opts.LabelOpts(font_weight='bold')),   #设置图例样式
        xaxis_opts=opts.AxisOpts(name="Year",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold')),   #x轴属性
        yaxis_opts=opts.AxisOpts(name="Percent",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold')),   #y轴属性
    )
    # 多个系列使用同一配置
    lineChart.set_series_opts(
        label_opts=opts.LabelOpts(formatter="{@[1]}%"),   #标签显示格式
        linestyle_opts=opts.LineStyleOpts( width=2)    #线宽
    )
    return lineChart

# 2-L 偿债能力表  (修改：将双Y轴改为1个)
liability = drawLineChart(finData,code,name+"偿债能力表",year_lst,["净负债率(%)_NetLiaRt","净资产负债率(%)_NetAstLiaRt","现金流动负债比_OpeCcurdb"])
# 3 -L   盈利能力表
profit = drawLineChart(finData,code,name+"盈利能力表",year_lst,["销售净利率(%)_Netprfrt","销售毛利率(%)_Gincmrt","营业利润率(%)_Opeprfrt"])
#4 - L 营运能力表
operate = drawLineChart(finData,code,name+"营运能力表",year_lst,["存货周转率(次)_Invtrtrrat","总资产周转率(次)_Totassrat"])




## 行业分析
def idstAna(df,idx,name,code):
    """
    指定某项index对行业进行分析
    name 本公司名
    code 本公司代码
    """
    comNumber = len(df)//5  #公司数量
    quartile = [int(i*comNumber) for i in [0.25,0.5,0.75]]
    idx_df = df[["comCd","comNm","endDt",idx]]    #只含该idx的df
    cpns = idx_df[idx_df.endDt==idx_df['endDt'].iloc[-1]].sort_values(idx).iloc[quartile]["comCd"].to_list()   #找出本年四分位公司代码

    this_idxs = idx_df[idx_df["comCd"]== code][idx].to_list()  #本公司指标
    this_idxs_change = [round((this_idxs[i]-this_idxs[i-1])*100/this_idxs[i-1],2) for i in range(1,len(this_idxs))]   #本公司变化率
    this_idxs_change.insert(0,0)   #第一年0
    idx_ch = ""
    for i in idx:
        if("\u4e00"<i<"\u9fa5"):
            idx_ch +=i

    # 净利率行业对比
    idxBar = (
        Bar(init_opts=opts.InitOpts(width="700px", height="400px",theme= ThemeType.LIGHT ))#设置图表画布宽度 
            .add_xaxis(list(df.endDt.apply(lambda x:str(x)[:4]).unique()))
            .extend_axis(yaxis=opts.AxisOpts())  #多添加一个y轴
            .add_yaxis(f"行业25%", idx_df[idx_df["comCd"]== cpns[2]][idx].to_list(),
                        label_opts=opts.LabelOpts(is_show=False) , yaxis_index=0, # 指定y轴，等于0时可以省略
                        itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][0],opacity=.8,),)  
            .add_yaxis(f"行业50%", idx_df[idx_df["comCd"]== cpns[1]][idx].to_list(),
                        label_opts=opts.LabelOpts(is_show=False), yaxis_index=0,
                        itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][1],opacity=.8,),)
            .add_yaxis(f"行业75%", idx_df[idx_df["comCd"]== cpns[0]][idx].to_list(),
                        label_opts=opts.LabelOpts(is_show=False), yaxis_index=0,
                        itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][2],opacity=.8,),)  
            .add_yaxis(name, this_idxs,
                        label_opts=opts.LabelOpts(formatter="{c}%"), yaxis_index=0,   
                        itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][3],opacity=.8,),
            )  

            .set_global_opts(
                title_opts=opts.TitleOpts(title=idx_ch+"\n行业对比"),
                legend_opts=opts.LegendOpts(textstyle_opts=opts.LabelOpts(font_weight='bold')),#设置图例属性
                #设置横纵坐标属性
                yaxis_opts=opts.AxisOpts(name="Percent",position="left",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold'),axislabel_opts=opts.LabelOpts(formatter="{value}%"))
            )
    )

    idxLine = (
        Line()
        .add_xaxis(list(df.endDt.apply(lambda x:str(x)[:4]).unique()))
        .add_yaxis(name+"增长率",this_idxs_change,
                label_opts=opts.LabelOpts(formatter="{@1}%"),yaxis_index=1, # 指定使用的Y轴
                 itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][19],opacity=1,),
                )  
        .set_global_opts(
            yaxis_opts=opts.AxisOpts(name=name+"增长率", position="right",
                                    name_textstyle_opts=opts.TextStyleOpts(font_weight='bold'),
                                    axislabel_opts=opts.LabelOpts(formatter="{value}%"))
        )
        .set_series_opts(
            label_opts=opts.LabelOpts(formatter=JsCode("function (params) {return params.value[1] + '%'}")),
            linestyle_opts=opts.LineStyleOpts( width=2),
        )
    )

    idxBar.overlap(idxLine)
    return idxBar


 #1 - R 行业净资产负债率
netAstLia = idstAna(finData,'净资产负债率(%)_NetAstLiaRt',name,code) 
#2 - R  行业销售净利率
netPro = idstAna(finData,"销售净利率(%)_Netprfrt",name,code)   
#3-R  行业资产周转率
assRat = idstAna(finData,"总资产周转率(次)_Totassrat",name,code)      



# 财务效率 行业状况
quartile = [int(i*comNumber) for i in [0.2,0.4,0.6,0.8]]

finEff_df = finData[["comCd","comNm","endDt",'销售净利率(%)_Netprfrt',"总资产周转率(次)_Totassrat","净资产负债率(%)_NetAstLiaRt"]][finData['endDt']==year_lst[-1]+"-12-31"]
finEff_df = pd.concat([finEff_df.sort_values("销售净利率(%)_Netprfrt").iloc[quartile],finEff_df[finEff_df.comCd==code]])   #只包含5家公司

# 净利率行业对比
finEff = (
    Bar(init_opts=opts.InitOpts(width="700px", height="400px",theme= ThemeType.LIGHT ))#设置图表画布宽度 
        .add_xaxis(finEff_df["comNm"].to_list(),)
        .extend_axis(yaxis=opts.AxisOpts()) #添加一个y轴
        .add_yaxis("总资产周转率", finEff_df["总资产周转率(次)_Totassrat"].to_list(),
                    label_opts=opts.LabelOpts(formatter="{@1}%") , yaxis_index=0,
                    itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][2],opacity=.7,),
                    
                    )  # 指定y轴，等于0时可以省  
        .set_global_opts(
            title_opts=opts.TitleOpts(title="财务效率"),
            legend_opts=opts.LegendOpts(textstyle_opts=opts.LabelOpts(font_weight='bold')),#设置图例属性
            #设置横纵坐标属性
            xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=-20)),
            yaxis_opts=opts.AxisOpts(name="Percent",position="left",name_textstyle_opts=opts.TextStyleOpts(font_weight='bold'),axislabel_opts=opts.LabelOpts(formatter="{value}%"))
        )
        .set_series_opts(
                    markpoint_opts=opts.MarkPointOpts(
                        data=[opts.MarkLineItem(name,coord=[finEff_df["comNm"].to_list()[-1],finEff_df["总资产周转率(次)_Totassrat"].to_list()[-1]])]
                    )
        )
)

finEffLine = (
    Line()
        .add_xaxis(finEff_df["comNm"].to_list(),)
        .add_yaxis("净资产负债率",finEff_df["净资产负债率(%)_NetAstLiaRt"].to_list(),
                label_opts=opts.LabelOpts(formatter="{@1}%"),yaxis_index=1,
                itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][1],opacity=1,),
                )  
        .add_yaxis("销售净利率",finEff_df["销售净利率(%)_Netprfrt"].to_list(),
                label_opts=opts.LabelOpts(is_show=False),yaxis_index=1,
                itemstyle_opts=opts.ItemStyleOpts(color=operate.options['color'][0],opacity=8,),
                )  
        .set_global_opts(
            yaxis_opts=opts.AxisOpts("", position="right",
                                    name_textstyle_opts=opts.TextStyleOpts(font_weight='bold'),
                                    axislabel_opts=opts.LabelOpts(formatter="{value}%")),
            xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=-20))
                        )
        .set_series_opts(
            label_opts=opts.LabelOpts(formatter=JsCode("function (params) {return params.value[1] + '%'}")),
            linestyle_opts=opts.LineStyleOpts( width=2),
        )
)
finEff.overlap(finEffLine)   #4-R
finEff.render_notebook()        

# 营收漏斗图 5 - C
incmOpe_df = finData[finData.endDt==finData['endDt'].iloc[-1]][["comCd","营业收入(元)_Incmope"]].sort_values("营业收入(元)_Incmope",ascending=False)  #当年排序后的营收
y = [incmOpe_df["营业收入(元)_Incmope"].rolling(comNumber//5).mean().iloc[i*comNumber//5-1] for i in range(1,6)]
x = [f"行业排名{i*20}%-{i*20+20}%\n平均营收=" for i in range(5)]
data = [[x[i]+str(y[i]//10000)+"w",i+1] for i in range(len(y))]

incmOpe =(
    Funnel(init_opts=opts.InitOpts(width="900px", height="500px"))
    .add(
        series_name="平均营收",
        data_pair=data,
        gap=1,
        tooltip_opts=opts.TooltipOpts(trigger="item", formatter="{a} <br/>{b}{c}"),
        sort_="ascending",  #正三角
        label_opts=opts.LabelOpts(is_show=True, position="inside"),
        itemstyle_opts=opts.ItemStyleOpts(border_color="#fff", border_width=1),
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(
            pos_top='30%',   #距上边位置
            title=f"{year_lst[-1]}行业营收分析", 
            subtitle=f"{name}={incmOpe_df[incmOpe_df.comCd == code]['营业收入(元)_Incmope'].iloc[-1]//10000}w"
            )
        )
)
incmOpe.render_notebook()

# 将上面定义好的图添加到 page
page = Page(layout=opts.PageLayoutOpts(justify_content='center', display="flex", flex_wrap="wrap"))    # 简单布局

page.add(
    expense,netAstLia,liability,netPro,profit,assRat,operate,finEff,
    incmOpe
)
page.render(f"./{name}公司财务指标分析表.html")