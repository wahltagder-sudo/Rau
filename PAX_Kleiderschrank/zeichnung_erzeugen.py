S=3.2  # px per cm
OX,OY=60,120
W,H=400,201
secs=[(0,100,"1 · 100 cm","Doppelhang"),(100,150,"2 · 50 cm","Fächer + Schuhe"),
      (150,250,"3 · 100 cm","Schubladen-Zentrale"),(250,300,"4 · 50 cm","Langhang"),
      (300,400,"5 · 100 cm","Hemden · Hosen · Wäsche")]
def X(c):return OX+c*S
def Y(c):return OY+c*S
out=[]
def r(x0,y0,x1,y1,**k):
    a=" ".join(f'{kk.replace("_","-")}="{v}"' for kk,v in k.items())
    out.append(f'<rect x="{X(x0):.1f}" y="{Y(y0):.1f}" width="{(x1-x0)*S:.1f}" height="{(y1-y0)*S:.1f}" {a}/>')
def t(x,y,s,size=13,anchor="middle",fill="#333",w="normal"):
    out.append(f'<text x="{X(x):.1f}" y="{Y(y):.1f}" font-size="{size}" text-anchor="{anchor}" fill="{fill}" font-weight="{w}">{s}</text>')
def shelf(a,b,y):r(a+1.8,y,b-1.8,y+1.8,fill="#e9e4dc",stroke="#b9b2a6",stroke_width=0.8)
def rail(a,b,y):
    out.append(f'<line x1="{X(a+3):.1f}" y1="{Y(y):.1f}" x2="{X(b-3):.1f}" y2="{Y(y):.1f}" stroke="#8d8d8d" stroke-width="4" stroke-linecap="round"/>')
def hangers(a,b,y,length,colors):
    n=len(colors);step=(b-a-8)/n
    for i,c in enumerate(colors):
        cx=a+4+step*(i+.5)
        out.append(f'<path d="M{X(cx):.1f},{Y(y):.1f} l0,{3*S:.1f} l{-step*S*0.45:.1f},{3*S:.1f} l{step*S*0.9:.1f},0 z" fill="none" stroke="#777" stroke-width="1"/>')
        r(cx-step*0.42,y+6,cx+step*0.42,y+length,fill=c,stroke="#00000022",rx=3)
def light(a,b,y):
    gid=f"g{a}_{y}"
    out.append(f'<defs><linearGradient id="{gid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff3c4" stop-opacity=".85"/><stop offset="1" stop-color="#fff3c4" stop-opacity="0"/></linearGradient></defs>')
    out.append(f'<polygon points="{X(a+4):.1f},{Y(y+1.8):.1f} {X(b-4):.1f},{Y(y+1.8):.1f} {X(b-1.8):.1f},{Y(y+45):.1f} {X(a+1.8):.1f},{Y(y+45):.1f}" fill="url(#{gid})"/>')
    r(a+4,y+1.8,b-4,y+2.6,fill="#fff7d6",stroke="#f2c84b",stroke_width=1)
def drawer(a,b,y0,y1,glass=False):
    r(a+2,y0+0.5,b-2,y1-0.5,fill="#cfe3ea" if glass else "#f7f5f1",stroke="#a9a39a",stroke_width=1,rx=2)
    if glass:
        for k in range(4):
            cx=a+10+k*(b-a-20)/3
            r(cx-6,y0+5,cx+6,y1-2,fill="#9aa9c7",rx=2)
    out.append(f'<line x1="{X(a+(b-a)/2-8):.1f}" y1="{Y(y0+3):.1f}" x2="{X(a+(b-a)/2+8):.1f}" y2="{Y(y0+3):.1f}" stroke="#555" stroke-width="3" stroke-linecap="round"/>')
def folded(a,b,y,n,color):
    w=(b-a-6)/n
    for i in range(n):
        for j in range(3):
            r(a+3+i*w+1,y-4.5*(j+1),a+3+(i+1)*w-1,y-4.5*j-0.4,fill=color,stroke="#00000030",rx=1.5)

out.append(f'<text x="{OX}" y="40" font-size="26" font-weight="bold" fill="#222">IKEA PAX · 400 × 58 × 201 cm · Innenansicht (Türen offen, Licht an)</text>')
out.append(f'<text x="{OX}" y="66" font-size="15" fill="#555">5 Korpusse: 100 | 50 | 100 | 50 | 100 · KOMPLEMENT-Innenausstattung · LED-Leisten mit Türsensor · Maße in cm, schematisch</text>')
# back wall
r(-3,-3,W+3,H+3,fill="#d8d3cb")
for a,b,_,_ in secs:
    r(a,0,b,H,fill="#fbfaf7",stroke="#9e978c",stroke_width=1.5)
    r(a,H-5,b,H,fill="#e2ddd4")  # Sockel
# section 1
shelf(0,100,30);light(0,100,30);rail(0,100,36)
hangers(0,100,36,62,["#3b5b8c","#e8e8e8","#87a6c9","#f0f0f0","#5d7da8","#c9d6e6","#e8e8e8","#2e4a6f"])
shelf(0,100,108);light(0,100,108);rail(0,100,114)
hangers(0,100,114,68,["#2f3542","#4a5568","#6b4f3a","#2f3542","#7a7f87","#1f2a3a"])
folded(0,100,29.5,3,"#c0a77f")
# section 2
for y in (30,58,86,114):shelf(100,150,y)
light(100,150,30)
folded(100,150,29.5,1,"#d9c3a5");folded(100,150,57.5,2,"#8fa3b8");folded(100,150,85.5,2,"#e6d5c3");folded(100,150,113.5,2,"#9b8070")
for y in (140,166):
    r(102,y,148,y+2,fill="#bdb6aa")
    for k in range(2):
        cx=108+k*20
        out.append(f'<ellipse cx="{X(cx+5):.1f}" cy="{Y(y-3):.1f}" rx="{9*S:.1f}" ry="{3.2*S:.1f}" fill="{["#5a3e2b","#ddd"][k]}" stroke="#0003"/>')
t(125,193,"Schuhauszüge",11)
# section 3
shelf(150,250,30);light(150,250,30);rail(150,250,36)
hangers(150,250,36,58,["#6d8b74","#e0c48d","#bba8d6","#d9d9d9","#a34a4a","#efe2cc","#4f6d7a"])
shelf(150,250,104);light(150,250,104)
folded(150,250,103.5,4,"#cfd8dc")
drawer(150,250,108,129,True);drawer(150,250,129,150,True);drawer(150,250,150,172);drawer(150,250,172,196)
folded(150,250,29.5,3,"#a7b6a3")
# section 4
shelf(250,300,30);light(250,300,30);rail(250,300,36)
hangers(250,300,36,150,["#3e3e3e","#7b5e57","#22303c"])
folded(250,300,29.5,1,"#b7a99a")
# section 5
shelf(300,400,30);light(300,400,30);rail(300,400,36)
hangers(300,400,36,72,["#ffffff","#dbe7f3","#f4f4f4","#c7d3e0","#ffffff","#e9eef5","#b8c7d9","#fafafa"])
shelf(300,400,118);light(300,400,118)
r(304,124,396,127,fill="#9a9a9a",rx=1)
for k in range(7):
    cx=310+k*12.5
    r(cx-3,127,cx+3,147,fill=["#2b2b2b","#5b6770","#8c7b6b","#2e3b4e","#6b6b6b","#3d4a3a","#4c4c4c"][k],rx=1)
t(350,150.5,"Hosenauszug",10)
for y0 in (152,174):
    r(304,y0,396,y0+21,fill="none",stroke="#999",stroke_width=1.2,rx=2)
    for k in range(1,12):
        out.append(f'<line x1="{X(304+k*7.6):.1f}" y1="{Y(y0):.1f}" x2="{X(304+k*7.6):.1f}" y2="{Y(y0+21):.1f}" stroke="#bbb" stroke-width="0.8"/>')
t(350,166,"Drahtkorb / Wäsche",11);t(350,188,"Drahtkorb / Sport",11)
folded(300,400,29.5,4,"#dcc9b0")
# labels
for a,b,n,d in secs:
    t((a+b)/2,H+11,n,14,w="bold");t((a+b)/2,H+17,d,13,fill="#555")
# dims
out.append(f'<line x1="{X(0)}" y1="{Y(-5.5)}" x2="{X(400)}" y2="{Y(-5.5)}" stroke="#444" marker-start="url(#ar)" marker-end="url(#ar)"/>')
t(200,-7.8,"400 cm (Korpusmaß, Wand netto ≥ 401 cm)",14,w="bold")
out.append(f'<line x1="{X(-8)}" y1="{Y(0)}" x2="{X(-8)}" y2="{Y(201)}" stroke="#444" marker-start="url(#ar)" marker-end="url(#ar)"/>')
out.append(f'<text transform="translate({X(-10)},{Y(100)}) rotate(-90)" font-size="14" text-anchor="middle" font-weight="bold">201 cm</text>')
interior="\n".join(out)

# closed view
out=[]
OY2=OY+H*S+150
def Y2(c):return OY2+c*S
out.append(f'<text x="{OX}" y="{OY2-25}" font-size="22" font-weight="bold" fill="#222">Außenansicht (Türen zu) · 8 Drehtüren mit Soft-Close · 1 Spiegeltür · Stangengriffe</text>')
out.append(f'<rect x="{X(-3)}" y="{Y2(-3)}" width="{406*S}" height="{207*S}" fill="#d8d3cb"/>')
out.append(f'<rect x="{X(0)}" y="{Y2(0)}" width="{400*S}" height="{201*S}" fill="#eeebe5" stroke="#9e978c"/>')
doors=[(0,50,"r"),(50,100,"l"),(100,150,"l"),(150,200,"r"),(200,250,"l"),(250,300,"m"),(300,350,"r"),(350,400,"l")]
for a,b,h in doors:
    fill="url(#mirror)" if h=="m" else "#f9f8f5"
    out.append(f'<rect x="{X(a+0.3):.1f}" y="{Y2(3):.1f}" width="{(b-a-0.6)*S:.1f}" height="{195*S:.1f}" fill="{fill}" stroke="#bdb6aa" stroke-width="1.2"/>')
    hx=b-4 if h in("r","m") else a+4
    if h=="m":hx=a+4
    out.append(f'<rect x="{X(hx)-3:.1f}" y="{Y2(85):.1f}" width="6" height="{30*S:.1f}" rx="3" fill="#2d2d2d"/>')
out.append(f'<rect x="{X(0)}" y="{Y2(198)}" width="{400*S}" height="{3*S}" fill="#cfc9bf"/>')
closed="\n".join(out)
TOTAL_H=int(OY2+201*S+60)
svg=f'''<svg xmlns="http://www.w3.org/2000/svg" width="{int(OX*2+400*S)}" height="{TOTAL_H}" font-family="Helvetica, Arial, sans-serif">
<defs><marker id="ar" viewBox="0 0 10 10" refX="5" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#444"/></marker>
<linearGradient id="mirror" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#e9f1f5"/><stop offset=".45" stop-color="#c6d5de"/><stop offset=".55" stop-color="#f4f8fa"/><stop offset="1" stop-color="#b9c9d3"/></linearGradient></defs>
<rect width="100%" height="100%" fill="#ffffff"/>
{interior}
{closed}
</svg>'''
open("/home/user/Rau/PAX_Kleiderschrank/PAX_400cm_Ansicht.svg","w").write(svg)
print(int(OX*2+400*S),TOTAL_H)
