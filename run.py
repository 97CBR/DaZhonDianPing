# -*- coding: utf-8 -*-
# @Time    : 11/28/2018 17:48
# @Author  : MARX·CBR
# @File    : run.py
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

        self.xls = xlrd.open_workbook('data.xlsx')
        self.readsheet = self.xls.sheets()[0]
        print(self.readsheet)
        # values = sheet.row_values(0)
        # print(values)
        self.cssfuntion=get_css()
        self.cssfuntion.add_value_in_dict()
        self.css=self.cssfuntion.mydict

    def get_string(self):
        for info in range(-3,149):
            values = self.readsheet.row_values(info)
            # print(values[-1])
            self.string = values[-1]
            self.tital=values[0]
            self.run_it()
            self.line+=1
        self.wbk.save('new.xls')
        pass

    # TODO:速度稍微有点慢，可以用多线程加速
    def run_it(self):

        try:
            soup=BeautifulSoup(self.string,'lxml')
            content = soup.findAll('span')
            for info in content:
                # print(info)
                c = info.get('class')
                # print(c[0])
                position = self.css[c[0]]
                position=str(position).replace('[','').replace(']','')
                position = position.split(',')
                # print(position)
                new = get_content().show(cx=position[0], cy=position[1])
                # print(new)
                rs = '<span class="{}"></span>'.format(c[0])
                # print(rs)
                self.string = self.string.replace(rs, new)
                # print(self.string)

            # v=eval(self.string)
            #
            # print(v[0])
            # for index,i in enumerate(v):
            #     soup=BeautifulSoup(v[index],'lxml')
            print(self.string)
            self.sheet.write(self.line, 1, self.string)
            self.sheet.write(self.line, 0, self.tital)
            return self.string
        except:
            pass

changeSentence().get_string()