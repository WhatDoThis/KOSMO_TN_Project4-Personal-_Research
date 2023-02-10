import pandas as pd, matplotlib.pyplot as plt, matplotlib.image as img, seaborn as sns

# 한글 안깨지게 함
from matplotlib import font_manager

font_path = "C:/Windows/Fonts/NGULIM.TTF"
font = font_manager.FontProperties(fname=font_path).get_name()

# 사이즈 및 배경 지정
fig = plt.figure(figsize=(17,9))
fig.canvas.set_window_title('2017.2 KETI THz sensor AZO/VOx Test')
fig.set_facecolor('white')

# 데이터 값 표시용 박스
box_fig = dict(boxstyle='square', facecolor='yellow')

#데이터 뽑아오기
df1 = pd.read_excel("Sample1.xlsx", engine="openpyxl")
sp1_A = df1['A']
sp1_V = df1['V']
sp1_R = df1['R']

df2 = pd.read_excel("Sample2.xlsx", engine="openpyxl")
sp2_A = df2['A']
sp2_V = df2['V']
sp2_R = df2['R']

df3 = pd.read_excel("Sample3.xlsx", engine="openpyxl")
sp3_A = df3['A']
sp3_V = df3['V']
sp3_R = df3['R']

result_A_li = list()
result_V_li = list()

# 결과2 값 도출과정
sp1_R_result_li = list(); sp2_R_result_li = list(); sp3_R_result_li = list()
sp1_V_result_li = list(); sp2_V_result_li = list(); sp3_V_result_li = list()

def getsp_R(w, x, y, z):
    for i in range(len(w)):
        if w[i] != 0 and w[i] == w.min():
            y.append(w[i-7]), y.append(w[i-6]), y.append(w[i-5]), y.append(w[i-4]), y.append(w[i-3]), y.append(w[i-2]), y.append(w[i-1]), y.append(w[i]), y.append(w[i+1]), y.append(w[i+2])
            z.append(x[i-7]), z.append(x[i-6]), z.append(x[i-5]), z.append(x[i-4]), z.append(x[i-3]), z.append(x[i-2]), z.append(x[i-1]), z.append(x[i]), z.append(x[i+1]), z.append(x[i+2])

getsp_R(sp1_R, sp1_V, sp1_R_result_li, sp1_V_result_li)
getsp_R(sp2_R, sp2_V, sp2_R_result_li, sp2_V_result_li)
getsp_R(sp3_R, sp3_V, sp3_R_result_li, sp3_V_result_li)

sp1_R_result = pd.Series(sp1_R_result_li); sp2_R_result = pd.Series(sp2_R_result_li); sp3_R_result = pd.Series(sp3_R_result_li)
sp1_V_result = pd.Series(sp1_V_result_li); sp2_V_result = pd.Series(sp2_V_result_li); sp3_V_result = pd.Series(sp3_V_result_li)


# 이미지 좌표 간결하게
img_x = [0]
img_y = [0]


# 그래프 그리는 메소드들
def pltApply(m1,m2,x,y,c,str,z,z2):
    ax = plt.subplot2grid((3, 4), (m1, m2), colspan=2)
    sc = plt.scatter(x, y, color=c)
    plt.plot(x, y, lw=0.7, color=c)
    plt.title("["+str+"]")
    plt.xlabel("전압(V)")
    plt.ylabel("측정 전류량(㎂)")
    plt.xticks(x, rotation=70, fontsize=8)
    plt.tick_params(axis='x', pad=6)
    plt.grid(ls='--', lw=0.5)
    for i in range(len(x)):
        if y[i] == y.max():
            result_A_li.append(y.max()) # 결과1용 데이터 뽑기
            result_V_li.append(x[i]) # 결과1용 데이터 뽑기
            plt.axvline(x=x[i], color='r', linestyle='dashed')
            plt.text(x[i], - z2, '%.1f' %x[i], ha='center', va='bottom', size=9, fontweight='bold')
            plt.text(x[i], y[i]+z, '%.1f' %y[i], ha='center', va='bottom', size=9, fontweight='bold', bbox=box_fig)
            plt.text(x[i+1], y[i+1]+z, '%.1f' %y[i+1], ha='center', va='bottom', size=9, fontweight='bold', bbox=box_fig)
            
    annot = ax.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points",
                    bbox=dict(boxstyle="round", fc="w"),
                    arrowprops=dict(arrowstyle="->"))
    annot.set_visible(False)

    def update_annot(xdata, ydata):
        annot.xy = (xdata,ydata)
        text = '%.1F' %xdata, '%.1f' %ydata
        annot.set_text(text)

    def hover(event):
        vis = annot.get_visible()
        if event.inaxes == ax:
            cont, ind = sc.contains(event)
            if cont:
                update_annot(event.xdata, event.ydata)
                annot.set_visible(True)
                fig.canvas.draw_idle()
            else:
                if vis:
                    annot.set_visible(False)
                    fig.canvas.draw_idle()
                    
    fig.canvas.mpl_connect("motion_notify_event", hover)


def resultApply(m1,m2,x,y,str,z):
    plt.subplot2grid((3, 4), (m1, m2), colspan=2)
    plt.scatter(x, y)
    plt.title("["+str+"]")
    plt.xlabel("전압(V)")
    plt.ylabel("측정 전류량(㎂)")
    plt.tick_params(axis='x', pad=6)
    plt.axhline(y=mean_A, color='r', linestyle='dashed')
    plt.axvline(x=mean_V, color='r', linestyle='dashed')
    plt.text(mean_V, mean_A+z, "평균값: V=" + '%.1f' %mean_V + ", A=" + '%.1f' %mean_A, ha='center', va='bottom', size=9, fontweight='bold', bbox=box_fig)
    for i in range(len(x)):
        plt.text(x[i], y[i]+z, "V=" + '%.1f' %x[i] + ", A=" + '%.1f' %y[i], ha='center', va='bottom', size=9, fontweight='bold', bbox=box_fig)
    
        
def result2Apply(x,y,c,z,s):
    plt.plot(x, y, lw=2.2, color=c, label=s)
    plt.legend(loc='upper left')
    for i in range(len(x)):
        if y[i] == y.min():
            plt.text(x[i], y[i]-z, '%.2f' %x[i] + ", A=" + '%.2f' %y[i], ha='center', va='bottom', size=9, fontweight='bold', bbox=box_fig)


sns.set_theme(font=font, style="darkgrid")

pltApply(0,0,sp1_V,sp1_A,"b","Sample1",3,5)
pltApply(1,0,sp2_V,sp2_A,"g","Sample2",0.7,-0.9)
pltApply(2,0,sp3_V,sp3_A,"cyan","Sample3",1.7,1.5)


result_V = pd.Series(result_V_li)
result_A = pd.Series(result_A_li)

mean_V = result_V.mean()
mean_A = result_A.mean()

resultApply(1,2,result_V,result_A,"[AZO/VOx박막 Test 결과 A:V]",1.7)


plt.subplot2grid((3, 4), (2, 2), colspan=2)
plt.title("[AZO/VOx박막 Test 결과 R:V 10-data Sampling]")
plt.xlabel("임의의 온도 증가량")
plt.ylabel("측정 저항값(㏀)")
result2Apply(sp1_V_result, sp1_R_result, "b", 36, "Sample1_R")
result2Apply(sp2_V_result, sp2_R_result, "g", 40, "Sample2_R")
result2Apply(sp3_V_result, sp3_R_result, "cyan", 35, "Sample3_R")

# 측정 과정에 대한 설명용
plt.subplot2grid((3, 4), (0, 2), colspan=1)
image = img.imread('Test.png')
plt.imshow(image)
plt.title("[측정방법]")
plt.xticks(img_x)
plt.yticks(img_y)

plt.figtext(.76,.87,"1. AZO란: Aluminum doped Zinc Oxide\n\n   사용처: 투명전극(반도체, 센서 등)", fontsize=11)
plt.figtext(.76,.78,"2. VOx란: Vanadium Oxide\n\n   사용처: 열 감지 센서 등", fontsize=11)
plt.figtext(.76,.71,"3. 결과: 열변화에 따른 VOx 의 저항변화를 알고 \n           싶었으나, 통전되어 확인이 어려웠음", fontsize=11)

plt.suptitle("AZO/VOx 박막(2/4um) Test 전압 대비 전류 변화", fontweight='bold', fontsize=15)

fig.tight_layout()

plt.show()