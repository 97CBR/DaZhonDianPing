# -*- coding: utf-8 -*-
# @Time    : 11/29/2018 14:33
# @Author  : MARX·CBR
# @File    : 修复无结构.py
import re
string="""['#无锡橙V my-80nO  员有特权#<br/>红豆 最  my-9scA 感觉 my-A1ce 火的\xa0 这家红豆 在 荟聚那 边 \xa0 my-fyio  my-MO6q  还 可以\xa0客流量也 不  my-lLMW \xa0现 在 貌似红豆 不 光卖衣服裤 my-aUEw 之 my-NY0v 的了\xa0 还 和京 my-oeCK  my-gVd9  my-g1Fp \xa0......', '红豆 的 衣 my-cbmT  很  my-9J38  的 ，穿起 my-76aa  很  my-FzOY  my-cbmT ， my-uVyC 别 是  内 衣 内 裤什 my-9nu6  的 贴 my-DIUD 衣 my-pb8M ，缓和透气，感 my-Qdgi  my-MxXt 质就 是  不  my-qAti ，布料相当好，跟 my-SpJc 友过 my-76aa  的 ，感 my-Qdgi  还  不  my-qAti ， my-SpJc 友 很 认 可 。这 个  店 面 不  my-3ojL ， my-2bQ2  my-rYGg  还  可 以。<br/> my-SpJc 友买咯几 个  内 衣 内 裤。赞 不  my-QFFL 口。下 次 再 my-76aa 。', ' my-Ogv5  my-6LLl  在 荟聚三楼，电 my-p1ya  my-MYDV ， my-Bbs0 方是 非  常  好  my-fScy  的 ， my-Ogv5  my-6LLl  在 锡城 my-3d2M  my-lyuS 也是 my-3xij  my-KIds 响亮 的  国  my-6nb2 品 my-SXvW ，打出 国 门， my-3d2M  my-lyuS  好  好  点  my-lsue 。商品还不 my-qAti ，买 的  my-3xij  my-KIds  多  的 是 my-Ogv5  my-6LLl 羽绒 my-cbmT  my-qyWz 内 my-zTH0 ，其他就 my-3xij  my-KIds 少 了 。羽绒 my-cbmT  my-qyWz 内 my-zTH0  my-mHQy 量 my-3xij  my-KIds  好 ，我 的 观 点 。可能有 点 偏颇。[微笑]', ' 这 家 店 的  my-Bbs0  my-Gxsj  是 在锡山区 的 东 my-BzWY 宜 家 荟 聚 购物 中  my-eHhe  的 2楼， my-Bbs0  my-Gxsj  my-TvYe 好找 的 ，店 家  的 抬 my-v1HL 字体非常 my-3ojL ，看 起  my-76aa  很 显眼，店里 my-Bbs0  my-Gxsj 也宽 my-tPlw ，环 my-rYGg  my-TvYe 好 的 ，灯光看 起  my-76aa 特别 的 明 my-Oxqh ，店 家  是 一 家  家 居专卖店， 这里 的 衣服品 my-mHQy 非常 的 不 my-qAti ，红豆也 是 一 个  my-2Nt1 牌 my-aUEw 知名度 my-3xij  my-KIds  my-Zan0 了， my-zMqA  my-We9m 内衣， my-GSnh 裤， my-qBWi 衣 my-1wAt  my-1wAt ， my-TlIU  my-LJjE 看 起  my-76aa 也 my-3xij  my-KIds  的 好，规 my-4fg1  很 齐 全 ，服务也 my-TvYe 好 的 ', ' my-Ogv5 豆是无 my-sPmE 本土的 my-cbmT 饰 品 牌， my-dSwp  my-NY0v  my-Tddh 店很多。<br/> 这一家 my-sYBu 子位于 my-sPmE 山区荟聚 my-r61q  my-pb8M 中心的三楼。<br/>在欧 my-Z8vz 超 my-d9Np  my-r61q  my-pb8M 出来必 my-zMqA  之 路 上，位置 可  以 。<br/>其 my-Tddh 面 不 宽， my-JAWK 呷哺呷哺垂直......', '这 里  的 红豆卖 的 是内衣和居 家 服。这 家 店 的 地方还 my-TvYe 大 的 ，衣服款式都出 my-Z2rb  了 ，要 my-w0Sj 购 什 么 my-TvYe 方 my-DYLi  的 。', '这家 店 在 宜 家荟 聚  的 二楼,靠近欧 my-Z8vz 超市 的 入 my-MYDV , 店 面 不  是  很 大, my-TvYe  my-tPlw  my-Oxqh  的 ,红 my-6LLl  的 衣服 品 质 还  是  不 错 的 ,也算 是  my-FWMD 锡 的 一 个  my-2Nt1 牌子了, 很 多连锁加 my-teLC  店 ,价格 还 行吧, 店 员态 my-Njsk  my-A1ce  好 , my-TvYe 热 my-bDbU  的 ']"""


from bs4 import BeautifulSoup
import lxml
from DaZhonDianPing.WebContent import get_content,get_css
import xlwt
import xlrd
class changeSentence:
    def __init__(self):
        self.string="""['店<span class="my-sk48"></span><span class="my-NrYN"></span>人推荐<span class="my-wj09"></span><span class="my-4jIa"></span>也吃<span class="my-GZtT"></span>无数<span class="my-TEU5"></span>，在荟<span class="my-NGRz"></span>三<span class="my-gzkn"></span>还<span class="my-sk48"></span>四<span class="my-gzkn"></span><span class="my-wj09"></span><span class="my-4jIa"></span>对于荟<span class="my-NGRz"></span><span class="my-wj09"></span><span class="my-Laz5"></span><span class="my-gzkn"></span>一<span class="my-BjCr"></span>没<span class="my-q24J"></span>一<span class="my-AhDb"></span><span class="my-OWQM"></span>晰<span class="my-wj09"></span>概念，总之<span class="my-ZOV6"></span><span class="my-sk48"></span>星<span class="my-ctu6"></span><span class="my-gXqZ"></span><span class="my-gzkn"></span>上<span class="my-ZOV6"></span><span class="my-sk48"></span><span class="my-GZtT"></span>，「南瓜羹」<span class="my-bc2W"></span><span class="my-AhDb"></span><span class="my-HSuQ"></span><span class="my-wj09"></span><span class="my-HSuQ"></span><span class="my-wj09"></span><span class="my-sk48"></span><span class="my-4jIa"></span><span class="my-wj09"></span><span class="my-ho4u"></span>爱，<span class="my-4jIa"></span>每<span class="my-TEU5"></span><span class="my-Rwik"></span><span class="my-aI3S"></span>起码......', '对<span class="my-8K0s"></span>一<span class="my-AhDb"></span>礼拜来<span class="my-cQfy"></span>三<span class="my-8lSV"></span>荟聚<span class="my-wj09"></span>我，居<span class="my-YnBj"></span>没吃<span class="my-Pyz5"></span><span class="my-bc2W"></span>家店<span class="my-q24J"></span>点说不<span class="my-Pyz5"></span>！<span class="my-8K0s"></span><span class="my-sk48"></span>周<span class="my-SiiR"></span><span class="my-xf5Y"></span>娃<span class="my-5Rno"></span>了<span class="my-AhDb"></span>双人餐来尝尝！<br/>韩<span class="my-Cv0W"></span><span class="my-kQTY"></span><span class="my-RLtg"></span><span class="my-wj09"></span>前菜，就<span class="my-sk48"></span>小碟子铺<span class="my-89pn"></span>一桌！还<span class="my-q24J"></span><span class="my-bc2W"></span><span class="my-AhDb"></span>送<span class="my-wj09"></span>烤土<span class="my-GZBY"></span>，宝贝之前看到......', '<span class="my-Vqyk"></span><span class="my-rp5n"></span><span class="my-Vqyk"></span>网红餐厅<span class="my-Vqyk"></span><span class="my-GZtT"></span><span class="my-6d6Z"></span>久<span class="my-GZtT"></span><br/>每<span class="my-TEU5"></span>到荟聚<span class="my-aI3S"></span><span class="my-CZ8w"></span><span class="my-oD5Z"></span><span class="my-vNYO"></span>来<span class="my-NCl8"></span>去<span class="my-6ge2"></span><br/>正好<span class="my-NCl8"></span>去<span class="my-BoaK"></span>家\xa0<span class="my-XTNq"></span><span class="my-oD5Z"></span><span class="my-CZ8w"></span>去<span class="my-BoaK"></span>家<span class="my-6ge2"></span>饭就<span class="my-CZ8w"></span><span class="my-vNYO"></span>来过来<span class="my-6ge2"></span>饭<span class="my-GZtT"></span><br/>在楼<span class="my-aanf"></span>买<span class="my-sx6E"></span>茶的时候<span class="my-Z0cM"></span><span class="my-QJ1Z"></span><span class="my-M77A"></span><span class="my-Vqyk"></span><span class="my-rp5n"></span><span class="my-Vqyk"></span>在哪里\xa0<span class="my-QJ1Z"></span>......', '好<span class="my-qP4L"></span><span class="my-sk48"></span><span class="my-mWMf"></span>三次<span class="my-6ge2"></span><br/><span class="my-mWMf"></span><span class="my-sd0J"></span>次<span class="my-6ge2"></span>中<span class="my-mQKc"></span><br/><span class="my-mWMf"></span><span class="my-WZOp"></span>次四<span class="my-kQub"></span><br/><span class="my-bc2W"></span><span class="my-AhDb"></span>能<span class="my-dzHf"></span>五<span class="my-kQub"></span><br/><span class="my-ULR2"></span>觉<span class="my-sd0J"></span>直有在改善<br/>菜<span class="my-0sha"></span><span class="my-v5VB"></span>来<span class="my-v5VB"></span>丰<span class="my-O5RZ"></span><br/><span class="my-6ge2"></span>法<span class="my-v5VB"></span>来<span class="my-v5VB"></span>多<br/>总<span class="my-sk48"></span>肉肉肉多无聊<br/>198套餐<span class="my-HSuQ"></span>的......', '<span class="my-6d6Z"></span><span class="my-3hgK"></span><span class="my-6ge2"></span><span class="my-wj09"></span><span class="my-1gso"></span><span class="my-WDNm"></span>\xa0五<span class="my-d81e"></span><span class="my-WDNm"></span>肥瘦均<span class="my-yzkx"></span>搭配<span class="my-1gso"></span><span class="my-WDNm"></span><span class="my-p2fr"></span>包生菜<span class="my-xPmV"></span>级<span class="my-3hgK"></span><span class="my-6ge2"></span>\xa0一<span class="my-Ilbl"></span>都不腻。<span class="my-8abE"></span><span class="my-m25b"></span>排<span class="my-wso8"></span>包裹着<span class="my-0gzm"></span><span class="my-0gzm"></span><span class="my-wj09"></span><span class="my-8abE"></span><span class="my-m25b"></span>\xa0搭配<span class="my-p2fr"></span>汁<span class="my-csEk"></span>道太棒了\xa0套餐里<span class="my-wj09"></span>部队<span class="my-bS6v"></span><span class="my-kQTY"></span>也<span class="my-6d6Z"></span><span class="my-VrHm"></span>\xa0香肠、鱼......', '<span class="my-FVdk"></span>店的<span class="my-dph1"></span>候<span class="my-iOaZ"></span>外<span class="my-q24J"></span>几<span class="my-FLSK"></span>在<span class="my-n8xM"></span>队，翻<span class="my-FLSK"></span><span class="my-6d6Z"></span>快，<span class="my-oD5Z"></span><span class="my-JDnW"></span><span class="my-n8xM"></span>多<span class="my-ncA1"></span>就<span class="my-VxJk"></span>吃<span class="my-aanf"></span>了。一坐下就<span class="my-q24J"></span>南瓜<span class="my-wY9N"></span><span class="my-7EE0"></span>来，紧接着服务<span class="my-M77A"></span><span class="my-Pyz5"></span>来烤<span class="my-5aTS"></span>司，点<span class="my-cQfy"></span>份肉就<span class="my-q24J"></span>主<span class="my-aAaw"></span>可<span class="my-RaLo"></span><span class="my-HEkf"></span><span class="my-PaZl"></span>挑，<span class="my-sFSg"></span><span class="my-q24J"></span>一份蔬<span class="my-P1PK"></span>拼<span class="my-ELv8"></span><span class="my-7EE0"></span>，幸亏看了<span class="my-6d6Z"></span>......', '今天<span class="my-mpkP"></span>午和<span class="my-7L75"></span><span class="my-Dl46"></span>來<span class="my-FVdk"></span>荟聚埸埸這<span class="my-cv8P"></span><span class="my-0Iy6"></span>樓<span class="my-wj09"></span>這家算韓式<span class="my-P1PK"></span><span class="my-wj09"></span><span class="my-pUmb"></span>廳吃午<span class="my-pUmb"></span>。<span class="my-mpkP"></span>午來<span class="my-FVdk"></span><span class="my-Dl46"></span><span class="my-7L75"></span><span class="my-oD5Z"></span><span class="my-wW5e"></span>多，<span class="my-oD5Z"></span><span class="my-JDnW"></span>等位子。今天在大眾點評<span class="my-kEPb"></span>買<span class="my-GZtT"></span>團購<span class="my-wj09"></span>198元<span class="my-wj09"></span>二<span class="my-Dl46"></span>套<span class="my-pUmb"></span>。<span class="my-0Iy6"></span>文<span class="my-Hc7n"></span><span class="my-sk48"></span><span class="my-aanf"></span><span class="my-P1PK"></span>送<span class="my-wj09"></span>，<span class="my-enrG"></span>台<span class="my-zHI3"></span><span class="my-q24J"></span><span class="my-sd0J"></span>份，但<span class="my-7R9w"></span>食<span class="my-sk48"></span><span class="my-CqcD"></span><span class="my-RaLo"></span>添加<span class="my-wj09"></span>。<span class="my-P1PK"></span><span class="my-wj09"></span><span class="my-csEk"></span>道<span class="my-kevq"></span><span class="my-sk48"></span><span class="my-VpnV"></span>本<span class="my-aanf"></span>帶辣<span class="my-wj09"></span>。', '朋友带<span class="my-4jIa"></span>来<span class="my-wj09"></span>，<span class="my-5Rno"></span><span class="my-wj09"></span>198<span class="my-wj09"></span><span class="my-Eo9n"></span><span class="my-pUmb"></span>，因<span class="my-ow2b"></span><span class="my-Eo9n"></span><span class="my-pUmb"></span>里东西很多，<span class="my-FLSK"></span>子<span class="my-aanf"></span><span class="my-aI3S"></span><span class="my-CRu1"></span><span class="my-oD5Z"></span>下<span class="my-eDd8"></span>！最后也<span class="my-u8W1"></span>得吃<span class="my-oD5Z"></span>下。<br/><span class="my-EKrn"></span><span class="my-5Rno"></span>，需<span class="my-NCl8"></span>自己捏，<span class="my-q24J"></span>海苔和<span class="my-z0hy"></span>枪鱼，但<span class="my-csEk"></span><span class="my-FeoM"></span>还<span class="my-oD5Z"></span>够。<br/>烤<span class="my-WDNm"></span>，<span class="my-Laz5"></span>种<span class="my-WDNm"></span><span class="my-aI3S"></span><span class="my-oD5Z"></span>错，......', '<span class="my-ZnAn"></span>意，<span class="my-ZnAn"></span>意，<span class="my-bc2W"></span>是<span class="my-sd0J"></span><span class="my-cWPK"></span><span class="my-CqcD"></span><span class="my-RaLo"></span>让你吃<span class="my-wj09"></span><span class="my-Chi7"></span><span class="my-iEDV"></span><span class="my-Chi7"></span><span class="my-iEDV"></span>饱还<span class="my-oD5Z"></span><span class="my-izZG"></span><span class="my-wj09"></span>韩式<span class="my-1gso"></span>肉<span class="my-QJ1Z"></span>。<br/>进<span class="my-QJ1Z"></span><span class="my-MID8"></span>下<span class="my-iSnV"></span>点完餐就送上<span class="my-GZtT"></span>人气考吐司三<span class="my-Y65R"></span>治，料<span class="my-CqcD"></span><span class="my-RaLo"></span>说是<span class="my-Chi7"></span><span class="my-iEDV"></span>多<span class="my-wj09"></span>，<span class="my-xPmV"></span>级<span class="my-mGO9"></span><span class="my-O5RZ"></span><span class="my-q24J"></span>没<span class="my-q24J"></span>，吐司火腿芝士<span class="my-sd0J"></span>个<span class="my-oD5Z"></span>......', '以<span class="my-QcVs"></span><span class="my-d9v8"></span>蝠里<span class="my-KN2M"></span>开了一<span class="my-cWPK"></span><span class="my-wj09"></span>，<span class="my-vmAO"></span>来<span class="my-oD5Z"></span>知为<span class="my-Ez3F"></span>关掉了。在<span class="my-d9v8"></span>蝠<span class="my-7gox"></span><span class="my-cWPK"></span>吃<span class="my-Pyz5"></span>两次，<span class="my-pXiL"></span><span class="my-8K0s"></span>芝士排骨实在<span class="my-t0pU"></span><span class="my-1rF4"></span>么爱，<span class="my-7gox"></span>时好<span class="my-qP4L"></span><span class="my-KN2M"></span><span class="my-t0pU"></span>烤肉、吐司之类<span class="my-wj09"></span>。饭<span class="my-5Rno"></span>还<span class="my-sk48"></span>服务员<span class="my-M53p"></span>忙捏<span class="my-wj09"></span>，感觉<span class="my-ho4u"></span>好吃<span class="my-wj09"></span>就<span class="my-sk48"></span>饭<span class="my-5Rno"></span>......']"""

        self.tital=""
        self.wbk = xlwt.Workbook()
        self.sheet = self.wbk.add_sheet(u'sheet1', cell_overwrite_ok=True)
        self.line = 0

        self.xls = xlrd.open_workbook('old.xlsx')
        self.readsheet = self.xls.sheets()[0]
        print(self.readsheet)
        # values = sheet.row_values(0)
        # print(values)
        self.cssfuntion=get_css()
        self.cssfuntion.add_value_in_dict()
        self.css=self.cssfuntion.mydict

        self.box_class = []


    def get_string(self):
        for info in range(0,515):
            values = self.readsheet.row_values(info)
            # print(values[-1])
            self.string = values[8]
            self.box_class = []
            try:
                for i in range(0, len(self.string) - 1):
                    if self.string[i] + self.string[i + 1] == 'bc':
                        # print(self.string[i:i + 6:])
                        self.box_class.append(self.string[i:i + 5:])
            except:
                ...
            self.tital=values[0]
            print(self.line,self.tital,self.string)
            self.run_it()
            print(self.line, self.tital, self.string)
            self.sheet.write(self.line, 10, self.string)
            self.sheet.write(self.line, 0, self.tital)
            self.line+=1
        self.wbk.save('deal.xls')
        pass

    # TODO:速度稍微有点慢，可以用多线程加速
    def run_it(self):
        try:
            # soup=BeautifulSoup(self.string,'lxml')
            content = self.box_class
            for info in content:
                # print(info)
                c = str(info)
                p=self.css.copy()
                # print(c[0])
                position = p[c]
                position=str(position).replace('[','').replace(']','')
                position = position.split(',')
                # print(position)
                new = get_content().show(cx=position[0], cy=position[1])
                # print(new)
                rs = '{}'.format(c)
                # print(rs)
                self.string = self.string.replace(rs, new)
                # print(self.string)

            # v=eval(self.string)
            #
            # print(v[0])
            # for index,i in enumerate(v):
            #     soup=BeautifulSoup(v[index],'lxml')
            print(self.tital, self.string)

            return self.string
        except:
            pass

changeSentence().get_string()