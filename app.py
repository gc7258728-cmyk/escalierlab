# -*- coding: utf-8 -*-
"""EscalierLab — dimensionnement géométrique pédagogique d'une volée droite.
Toutes les fonctions, ressources et présentations sont réunies dans ce fichier.
"""
from __future__ import annotations

import base64
import copy
import hashlib
import html
import io
import json
import math
import re
import uuid
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab import rl_config
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Flowable, KeepTogether

VERSION = '1.0'
WINE, COPPER, INK, CREAM = '#702D46', '#A95830', '#302733', '#FAF5ED'
LAWS = ['Fixe', 'Uniforme', 'Triangulaire']
KINDS = ['Hypothèse de projet', 'Dimensions relevées', 'Dimensions issues de plans']
LIMITS = {'H': (.30, 12.), 'R': (.30, 20.), 'W': (.30, 6.), 'n': (2, 60)}
CRITERIA = dict(h_min=14., h_max=18., g_min=28., g_max=32., b_min=60., b_max=64., angle_max=40., width_min=.80)
EPS = 1e-7
REFS = [
    ('Proportions des escaliers et relation de Blondel · guide français', 'https://www.ecologie.gouv.fr/sites/default/files/publications/2019%2007%20guide_DHUP_erp-existants.pdf?lang=en'),
    ('Dépendances Streamlit Community Cloud', 'https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies')]


def uid(): return uuid.uuid4().hex[:12]


def fmt(value, digits=2):
    return f'{float(value):,.{digits}f}'.replace(',', ' ').replace('.', ',')


def num(value, label, low, high, integer=False):
    if isinstance(value, (bool, np.bool_)): raise ValueError(label+' : un nombre est attendu.')
    try: x = float(value.replace(',', '.').replace(' ', '')) if isinstance(value, str) else float(value)
    except (ValueError, TypeError): raise ValueError(label+' : nombre invalide.') from None
    if not math.isfinite(x) or not low <= x <= high or (integer and not x.is_integer()):
        raise ValueError(f'{label} : valeur requise entre {low} et {high}'+(' (entière).' if integer else '.'))
    return int(x) if integer else x


def text(value, label, limit=800, required=False):
    if not isinstance(value, str) or len(value)>limit or (required and not value.strip()): raise ValueError(label+' : texte absent ou trop long.')
    return value.strip()


def fixed_uncertainty():
    return {v: dict(law='Fixe', delta_cm=0.) for v in ['H','R','W']}


def new_project(demo=False):
    return dict(app='EscalierLab', version=VERSION, id=uid(), name='Volée droite · exemple fictif' if demo else 'Mon escalier',
                demo=demo, H=3., R=4.76, W=1., n=18, kind=KINDS[0],
                source='Géométrie entièrement fictive : H=3 m ; R=4,76 m ; W=1 m ; 18 hauteurs.' if demo else '',
                notes='', criteria=copy.deepcopy(CRITERIA),
                criteria_source='Critères pédagogiques d’exercice, à adapter au projet ; aucune norme locale affirmée.',
                uncertainty=dict(H=dict(law='Fixe',delta_cm=0.), R=dict(law='Uniforme',delta_cm=4.), W=dict(law='Fixe',delta_cm=0.)) if demo else fixed_uncertainty(),
                measures=dict(source='', values=[], spread_cm=.5, closure_cm=.5))


def validate(p):
    if not isinstance(p, dict) or p.get('app')!='EscalierLab' or p.get('version')!=VERSION:
        raise ValueError('Sauvegarde EscalierLab 1.0 attendue.')
    q=copy.deepcopy(p)
    for key,limit,required in [('id',50,True),('name',120,True),('source',800,False),('notes',1200,False),('criteria_source',800,True)]:
        q[key]=text(q.get(key),key,limit,required)
    if type(q.get('demo')) is not bool: raise ValueError('Indicateur fictif invalide.')
    if q.get('kind') not in KINDS: raise ValueError('Nature des dimensions inconnue.')
    for key,(low,high) in LIMITS.items(): q[key]=num(q.get(key),key,low,high,key=='n')
    c=q.get('criteria')
    if not isinstance(c,dict) or set(c)!=set(CRITERIA): raise ValueError('Critères incomplets.')
    for key in ['h_min','h_max','g_min','g_max','b_min','b_max']:
        c[key]=num(c[key],key,.1,200.)
    c['angle_max']=num(c['angle_max'],'Angle maximal',1.,89.)
    c['width_min']=num(c['width_min'],'Largeur minimale',.1,6.)
    for a,b in [('h_min','h_max'),('g_min','g_max'),('b_min','b_max')]:
        if c[a]>=c[b]: raise ValueError('Chaque minimum de critère doit être strictement inférieur au maximum.')
    u=q.get('uncertainty')
    if not isinstance(u,dict) or set(u)!=set(['H','R','W']): raise ValueError('Hypothèses d’incertitude incomplètes.')
    for key in ['H','R','W']:
        if not isinstance(u[key],dict) or u[key].get('law') not in LAWS: raise ValueError('Loi inconnue pour '+key)
        delta=num(u[key].get('delta_cm'),'Amplitude '+key,0.,1000.)
        if u[key]['law']=='Fixe': delta=0.
        if q[key]-delta/100<=0: raise ValueError('La borne basse de '+key+' doit rester strictement positive.')
        if q[key]+delta/100>LIMITS[key][1]: raise ValueError('La borne haute de '+key+' dépasse la limite du modèle.')
        u[key]=dict(law=u[key]['law'],delta_cm=delta)
    m=q.get('measures')
    if not isinstance(m,dict) or not isinstance(m.get('values'),list) or len(m['values'])>60: raise ValueError('60 relevés maximum.')
    m['source']=text(m.get('source'),'Source des relevés',800)
    m['values']=[num(x,'Hauteur relevée (cm)',.01,100.) for x in m['values']]
    m['spread_cm']=num(m.get('spread_cm'),'Écart maximal entre marches (cm)',.01,20.)
    m['closure_cm']=num(m.get('closure_cm'),'Tolérance de fermeture (cm)',.001,20.)
    if m['values'] and not m['source']: raise ValueError('Indiquer une source pour les hauteurs relevées.')
    return q


def geometry(H,R,W,n):
    h=H/n;g=R/(n-1)
    return dict(h_cm=100*h,g_cm=100*g,blondel_cm=100*(2*h+g),angle_deg=np.degrees(np.arctan(h/g)),
                footprint_m2=R*W,n_treads=n-1)


def conditions(p, values=None):
    d=values or geometry(p['H'],p['R'],p['W'],p['n']);c=p['criteria']
    width=p['W'] if values is None else values['width_m']
    return {
        'hauteur_ok':(d['h_cm']>=c['h_min']-EPS)&(d['h_cm']<=c['h_max']+EPS),
        'giron_ok':(d['g_cm']>=c['g_min']-EPS)&(d['g_cm']<=c['g_max']+EPS),
        'blondel_ok':(d['blondel_cm']>=c['b_min']-EPS)&(d['blondel_cm']<=c['b_max']+EPS),
        'pente_ok':d['angle_deg']<=c['angle_max']+EPS,
        'largeur_ok':width>=c['width_min']-EPS/100}


def diagnostic(p):
    d=geometry(p['H'],p['R'],p['W'],p['n']);c=p['criteria'];ok=conditions(p)
    details=[('Hauteur de marche',fmt(d['h_cm'])+' cm',f"{fmt(c['h_min'])} à {fmt(c['h_max'])} cm",'hauteur_ok'),
             ('Giron',fmt(d['g_cm'])+' cm',f"{fmt(c['g_min'])} à {fmt(c['g_max'])} cm",'giron_ok'),
             ('Relation 2h + g',fmt(d['blondel_cm'])+' cm',f"{fmt(c['b_min'])} à {fmt(c['b_max'])} cm",'blondel_ok'),
             ('Pente de foulée',fmt(d['angle_deg'])+'°','≤ '+fmt(c['angle_max'])+'°','pente_ok'),
             ('Largeur utile',fmt(p['W'])+' m','≥ '+fmt(c['width_min'])+' m','largeur_ok')]
    return pd.DataFrame([dict(controle=label,valeur=value,critere=limit,etat='À documenter' if not p['source'] else 'Dans le critère' if bool(ok[key]) else 'Écart au critère') for label,value,limit,key in details])


def candidates(p,lo=10,hi=30):
    lo=num(lo,'Minimum de hauteurs',2,60,True);hi=num(hi,'Maximum de hauteurs',lo,60,True)
    rows=[];c=p['criteria'];target=(c['b_min']+c['b_max'])/2
    for n in range(lo,hi+1):
        q=dict(p,n=n);d=geometry(q['H'],q['R'],q['W'],n);ok=conditions(q)
        rows.append(dict(hauteurs=n,girons=n-1,hauteur_cm=d['h_cm'],giron_cm=d['g_cm'],blondel_cm=d['blondel_cm'],
                         pente_deg=d['angle_deg'],respect_geometrique=all(bool(x) for x in ok.values()),distance_pas_cible_cm=abs(d['blondel_cm']-target)))
    return pd.DataFrame(rows)


def measure_stats(p):
    m=p['measures'];x=np.asarray(m['values'],dtype=float)
    if not len(x): return None
    full=len(x)==p['n'];gap=float(x.sum()-100*p['H']) if full else None
    return dict(count=len(x),complete=full,mean=float(x.mean()),std=float(x.std(ddof=1)) if len(x)>1 else None,
                minimum=float(x.min()),maximum=float(x.max()),spread=float(np.ptp(x)),closure=gap,
                spread_ok=bool(np.ptp(x)<=m['spread_cm']+EPS),closure_ok=bool(abs(gap)<=m['closure_cm']+EPS) if full else None)


def fingerprint(p): return hashlib.sha256(json.dumps(p,ensure_ascii=False,sort_keys=True,allow_nan=False).encode()).hexdigest()


@st.cache_data(ttl=1800,max_entries=4,show_spinner=False)
def simulate(serialized,count=10000,seed=42):
    p=validate(json.loads(serialized))
    if not p['source']: raise ValueError('Documenter les dimensions ou déclarer explicitement leur caractère hypothétique avant de simuler.')
    count=num(count,'Tirages',1000,100000,True);seed=num(seed,'Graine',0,2**32-1,True)
    rng=np.random.default_rng(seed);inputs={}
    for key in ['H','R','W']:
        ref=p[key];u=p['uncertainty'][key];delta=u['delta_cm']/100
        inputs[key]=np.full(count,ref) if u['law']=='Fixe' or delta==0 else rng.uniform(ref-delta,ref+delta,count) if u['law']=='Uniforme' else rng.triangular(ref-delta,ref,ref+delta,count)
    d=geometry(inputs['H'],inputs['R'],inputs['W'],p['n']);d['width_m']=inputs['W']
    tests=conditions(p,d)
    df=pd.DataFrame(dict(H_m=inputs['H'],R_m=inputs['R'],W_m=inputs['W'],h_cm=d['h_cm'],g_cm=d['g_cm'],
                         blondel_cm=d['blondel_cm'],pente_deg=d['angle_deg'],emprise_m2=d['footprint_m2'],**tests))
    df['tous_criteres']=df[list(tests)].all(axis=1)
    return df


def wilson(k,n):
    if not 0<=k<=n or n<=0: raise ValueError('Comptages invalides.')
    z=1.95996398454;prob=k/n;den=1+z*z/n;center=(prob+z*z/(2*n))/den
    half=z*math.sqrt(prob*(1-prob)/n+z*z/(4*n*n))/den
    return max(0.,center-half),min(1.,center+half)


def csv_bytes(frame):
    safe=frame.copy()
    for col in safe.select_dtypes(include=['object','string']).columns:
        safe[col]=safe[col].map(lambda x:"'"+x if isinstance(x,str) and x.lstrip().startswith(('=','+','-','@')) else x)
    return safe.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig')


def measures_frame(p):
    return pd.DataFrame(dict(numero=range(1,len(p['measures']['values'])+1),hauteur_cm=p['measures']['values']))


def import_measures(raw,filename):
    if len(raw)>1_000_000: raise ValueError('Fichier limité à 1 Mo.')
    if filename.lower().endswith('.csv'):
        df=pd.read_csv(io.BytesIO(raw),sep=';',encoding='utf-8-sig',dtype=str,keep_default_na=False)
    elif filename.lower().endswith('.xlsx'):
        import zipfile
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            if sum(x.file_size for x in z.infolist())>10_000_000: raise ValueError('Classeur décompressé trop volumineux.')
        df=pd.read_excel(io.BytesIO(raw),sheet_name='Releves',dtype=str,keep_default_na=False,engine='openpyxl')
    else: raise ValueError('CSV ou XLSX attendu.')
    if set(df.columns)!=set(['numero','hauteur_cm']) or not 1<=len(df)<=60:
        raise ValueError('Colonnes numero et hauteur_cm requises ; entre 1 et 60 lignes.')
    order=[num(x,'Numéro',1,60,True) for x in df.numero]
    if order!=list(range(1,len(df)+1)): raise ValueError('Numéros attendus dans l’ordre, de 1 à n, sans doublon.')
    return [num(x,'Hauteur relevée (cm)',.01,100.) for x in df.hauteur_cm]


def backup_bytes(p): return json.dumps(validate(p),ensure_ascii=False,indent=2,allow_nan=False).encode('utf-8')


def restore(raw):
    if len(raw)>100_000: raise ValueError('Sauvegarde limitée à 100 ko.')
    return validate(json.loads(raw.decode('utf-8-sig')))


def profile_points(p):
    h=p['H']/p['n'];g=p['R']/(p['n']-1);x=[0.];y=[0.]
    for i in range(p['n']):
        x.append(i*g);y.append((i+1)*h)
        if i<p['n']-1:x.append((i+1)*g);y.append((i+1)*h)
    return np.asarray(x),np.asarray(y)


def profile_figure(p):
    x,y=profile_points(p);d=geometry(p['H'],p['R'],p['W'],p['n']);fig=go.Figure()
    fig.add_scatter(x=x,y=y,mode='lines',line=dict(color=WINE,width=3),name='Profil des marches',hovertemplate='x=%{x:.3f} m<br>z=%{y:.3f} m<extra>Profil</extra>')
    extension=min(.7,p['R']/4)
    fig.add_scatter(x=[-extension,0,None,p['R'],p['R']+extension],y=[0,0,None,p['H'],p['H']],mode='lines',line=dict(color=INK,width=5),name='Sols finis')
    fig.add_scatter(x=[0,p['R']],y=[p['H']/p['n'],p['H']],mode='lines',line=dict(color=COPPER,dash='dot',width=2),name='Ligne des nez de marche')
    fig.update_xaxes(title='Reculement horizontal · m',range=[-extension,p['R']+extension])
    fig.update_yaxes(title='Altitude · m',scaleanchor='x',scaleratio=1,range=[-p['H']*.08,p['H']*1.13])
    fig.add_annotation(x=0,y=p['H']*1.09,text=f"H = {fmt(p['H'])} m · {p['n']} hauteurs",showarrow=False,xanchor='left',font=dict(color=INK,size=11))
    return fig


def plan_figure(p):
    fig=go.Figure();n=p['n']-1;g=p['R']/n
    fig.add_scatter(x=[0,p['R'],p['R'],0,0],y=[0,0,p['W'],p['W'],0],fill='toself',fillcolor='#EEDCD3',line=dict(color=WINE,width=2),mode='lines',name='Emprise de la volée')
    xs=[];ys=[]
    for i in range(1,n):xs.extend([i*g,i*g,None]);ys.extend([0,p['W'],None])
    fig.add_scatter(x=xs,y=ys,mode='lines',line=dict(color=COPPER,width=1),name='Limites des girons')
    fig.add_annotation(x=p['R']*.80,y=p['W']/2,ax=p['R']*.20,ay=p['W']/2,axref='x',ayref='y',text='Montée',showarrow=True,arrowhead=2,arrowcolor=WINE,font=dict(color=WINE),arrowsize=1.5)
    fig.update_xaxes(title='Reculement R · m',range=[-.05*p['R'],1.05*p['R']]);fig.update_yaxes(title='Largeur utile W · m',scaleanchor='x',range=[-.25*p['W'],1.4*p['W']])
    return fig


def volume_figure(p):
    n=p['n'];h=p['H']/n;g=p['R']/(n-1);w=p['W'];vx=[];vy=[];vz=[];ii=[];jj=[];kk=[];colors_faces=[]
    def face(points,color):
        start=len(vx)
        for x,y,z in points:vx.append(x);vy.append(y);vz.append(z)
        ii.extend([start,start]);jj.extend([start+1,start+2]);kk.extend([start+2,start+3]);colors_faces.extend([color,color])
    for j in range(n):
        x=j*g;z=j*h
        face([(x,0,z),(x,w,z),(x,w,z+h),(x,0,z+h)],'#9D5B70')
        if j<n-1:face([(x,0,z+h),(x+g,0,z+h),(x+g,w,z+h),(x,w,z+h)],'#E1BFA8' if j%2 else '#EED6C0')
    fig=go.Figure(go.Mesh3d(x=vx,y=vy,z=vz,i=ii,j=jj,k=kk,facecolor=colors_faces,flatshading=True,hoverinfo='skip',lighting=dict(ambient=.8,diffuse=.7)))
    scale=max(p['R'],p['W'],p['H'])
    fig.update_layout(scene=dict(xaxis=dict(visible=False),yaxis=dict(visible=False),zaxis=dict(visible=False),
                                aspectmode='manual',aspectratio=dict(x=p['R']/scale,y=p['W']/scale,z=p['H']/scale),
                                camera=dict(eye=dict(x=-.9,y=-1.1,z=1.5))))
    return fig


def excel_bytes(p,sim=None):
    wb=Workbook();wb.remove(wb.active);d=geometry(p['H'],p['R'],p['W'],p['n'])
    tables={
        'Projet':pd.DataFrame([('Nom',p['name']),('Exemple fictif',str(p['demo'])),('Nature',p['kind']),('Source des dimensions',p['source']),('H_m',p['H']),('R_m',p['R']),('W_m',p['W']),('n',p['n']),('Source des critères',p['criteria_source']),('Notes',p['notes'])],columns=['champ','valeur']),
        'Diagnostic':diagnostic(p),'Comparaison':candidates(p,2,60),'Releves':measures_frame(p),
        'Hypotheses':pd.DataFrame([dict(dimension=k,reference_m=p[k],loi=p['uncertainty'][k]['law'],amplitude_cm=p['uncertainty'][k]['delta_cm']) for k in ['H','R','W']]),
        'Source_releves':pd.DataFrame([dict(source=p['measures']['source'],ecart_max_cm=p['measures']['spread_cm'],fermeture_max_cm=p['measures']['closure_cm'])]),
        'Simulation':pd.DataFrame([dict(etat='Aucune simulation actuelle')]) if sim is None else pd.DataFrame([dict(tirages=sim['count'],graine=sim['seed'],hauteurs_fixes=p['n'],probabilite_tous=float(sim['df'].tous_criteres.mean()),P05_giron_cm=sim['df'].g_cm.quantile(.05),mediane_giron_cm=sim['df'].g_cm.median(),P95_giron_cm=sim['df'].g_cm.quantile(.95))])}
    for name,frame in tables.items():
        ws=wb.create_sheet(name);ws.append(list(frame.columns))
        for row in frame.itertuples(index=False,name=None):ws.append([x.item() if isinstance(x,np.generic) else None if pd.isna(x) else x for x in row])
        for row in ws:
            for cell in row:
                if isinstance(cell.value,str):cell.data_type='s'
                cell.alignment=Alignment(wrap_text=True,vertical='top')
        for cell in ws[1]:cell.font=Font(color='FFFFFF',bold=True);cell.fill=PatternFill('solid',fgColor='702D46')
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=28 if name!='Projet' else 48
    out=io.BytesIO();wb.save(out);return out.getvalue()


class StairDrawing(Flowable):
    def __init__(self,p):
        Flowable.__init__(self);self.p=p;self.width=17.5*cm;self.height=7.3*cm
    def draw(self):
        p=self.p;c=self.canv;bottom=38;scale=min((self.width-100)/p['R'],(self.height-75)/p['H']);left=(self.width-p['R']*scale)/2;x,y=profile_points(p)
        c.saveState();c.setFillColor(colors.HexColor(CREAM));c.roundRect(0,0,self.width,self.height,8,fill=1,stroke=0)
        path=c.beginPath();path.moveTo(left+x[0]*scale,bottom+y[0]*scale)
        for a,b in zip(x[1:],y[1:]):path.lineTo(left+a*scale,bottom+b*scale)
        c.setStrokeColor(colors.HexColor(WINE));c.setLineWidth(1.4);c.drawPath(path)
        c.setLineWidth(.7);c.setDash(3,2);c.setStrokeColor(colors.HexColor(COPPER));c.line(left,bottom+p['H']/p['n']*scale,left+p['R']*scale,bottom+p['H']*scale);c.setDash()
        c.setStrokeColor(colors.HexColor(INK));c.setLineWidth(.5)
        c.line(left,bottom-10,left+p['R']*scale,bottom-10);c.line(left-10,bottom,left-10,bottom+p['H']*scale)
        c.setFont('StairVera',8);c.setFillColor(colors.HexColor(INK));c.drawString(left,bottom-25,'R = '+fmt(p['R'])+' m ; largeur '+fmt(p['W'])+' m')
        c.saveState();c.translate(left-20,bottom);c.rotate(90);c.drawString(0,0,'H = '+fmt(p['H'])+' m');c.restoreState()
        c.setFont('StairVera',7);c.drawString(14,self.height-15,'Profil de la volée · sols finis · paliers exclus · dimensions géométriques')
        c.restoreState()


def pdf_bytes(p,sim=None):
    if 'StairVera' not in pdfmetrics.getRegisteredFontNames():
        fonts=Path(rl_config.__file__).parent/'fonts'
        pdfmetrics.registerFont(TTFont('StairVera',str(fonts/'Vera.ttf')));pdfmetrics.registerFont(TTFont('StairVeraBold',str(fonts/'VeraBd.ttf')))
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle('BodyStair',fontName='StairVera',fontSize=9,leading=13,spaceAfter=7,splitLongWords=True))
    styles.add(ParagraphStyle('TitleStair',parent=styles['BodyStair'],fontName='StairVeraBold',fontSize=22,leading=28,textColor=colors.HexColor(WINE),spaceAfter=13))
    styles.add(ParagraphStyle('HeadStair',parent=styles['BodyStair'],fontName='StairVeraBold',fontSize=12,leading=17,spaceBefore=12,keepWithNext=True))
    styles.add(ParagraphStyle('CellStair',parent=styles['BodyStair'],fontSize=8,leading=11,spaceAfter=0))
    def par(t,style='BodyStair'):return Paragraph(html.escape(str(t)).replace('\n','<br/>'),styles[style])
    def table(headers,rows,widths):
        t=Table([[par(x,'CellStair') for x in headers]]+[[par(x,'CellStair') for x in row] for row in rows],colWidths=widths,repeatRows=1,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#ECD3C5')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor(CREAM)]),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]));return t
    out=io.BytesIO();doc=SimpleDocTemplate(out,pagesize=(21*cm,29.7*cm),leftMargin=1.7*cm,rightMargin=1.7*cm,topMargin=1.6*cm,bottomMargin=1.7*cm)
    d=geometry(p['H'],p['R'],p['W'],p['n'])
    story=[par('EscalierLab · Note géométrique','TitleStair'),par(('EXEMPLE FICTIF · ' if p['demo'] else '')+p['name']),
           par('Une volée droite à marches uniformes ; le sol fini supérieur sert d’arrivée. Géométrie selon les critères du projet, sans dimensionnement du ferraillage ou de la résistance.'),
           par(f"Dimensions : H={fmt(p['H'])} m ; R={fmt(p['R'])} m ; W={fmt(p['W'])} m ; {p['n']} hauteurs et {p['n']-1} girons."),
           par('Nature : '+p['kind']+' ; source : '+(p['source'] or 'à documenter')),StairDrawing(p),
           par('Calculs et contrôle des critères','HeadStair'),par('h = H/n ; g = R/(n-1) ; pas = 2h+g ; pente de foulée = atan(h/g). Emprise R×W, paliers et trémie exclus.'),
           table(['Contrôle','Valeur','Critère saisi','État'],diagnostic(p).values.tolist(),[4.3*cm,3.3*cm,5.3*cm,4.7*cm]),
           par('Origine des critères : '+p['criteria_source']),par('Emprise de la volée : '+fmt(d['footprint_m2'])+' m². Le seul contrôle de 2h+g ne vérifie ni largeur, ni pente, ni échappée.')]
    feasible=candidates(p,2,60);feasible=feasible[feasible.respect_geometrique].sort_values('distance_pas_cible_cm')
    story.append(par('Autres configurations géométriques','HeadStair'))
    if len(feasible):
        rows=[[int(r.hauteurs),fmt(r.hauteur_cm),fmt(r.giron_cm),fmt(r.blondel_cm)] for r in feasible.head(8).itertuples()]
        story.append(table(['Hauteurs','h (cm)','g (cm)','2h+g (cm)'],rows,[3.3*cm,4.3*cm,4.3*cm,5.7*cm]))
        story.append(par('Au plus huit configurations affichées, classées par proximité de la cible (minimum Blondel + maximum Blondel)/2. Ce classement n’est pas une optimisation structurelle.'))
    else:story.append(par('Aucune configuration entre 2 et 60 hauteurs ne respecte tous les critères saisis pour ces dimensions.'))
    ms=measure_stats(p)
    if ms:
        story+=[par('Relevés de hauteurs distincts des simulations','HeadStair'),par('Source : '+p['measures']['source']),
                par(f"{ms['count']} relevés ; moyenne {fmt(ms['mean'],4)} cm ; écart min–max {fmt(ms['spread'],4)} cm ; écart-type échantillonnal "+(fmt(ms['std'],4)+' cm.' if ms['std'] is not None else 'non défini (n<2).')),
                par('Fermeture : '+(fmt(ms['closure'],4)+' cm (somme des hauteurs moins H).' if ms['complete'] else 'non évaluée : le nombre de relevés diffère du nombre de hauteurs.'))]
    story.append(par('Hypothèses des dimensions variables','HeadStair'))
    story.append(table(['Dimension','Référence (m)','Loi','Amplitude ± (cm)'],[[k,fmt(p[k],3),p['uncertainty'][k]['law'],fmt(p['uncertainty'][k]['delta_cm'],3)] for k in ['H','R','W']],[3.3*cm,4.3*cm,4.3*cm,5.7*cm]))
    if sim:
        frame=sim['df'];prob=float(frame.tous_criteres.mean());lo,hi=wilson(int(frame.tous_criteres.sum()),len(frame))
        story += [par('Résultat Monte-Carlo actuel','HeadStair'),par(f"{sim['count']} tirages ; graine {sim['seed']} ; nombre de hauteurs fixé à {p['n']}. Respect simultané des critères : {fmt(100*prob,3)} %. Intervalle Wilson 95 % de la fréquence : {fmt(100*lo,3)} à {fmt(100*hi,3)} %."),
                  par('Wilson mesure la précision numérique de la simulation, pas la validité des hypothèses. Les quantiles décrivent les géométries simulées.'),
                  table(['Grandeur','P05','Médiane','P95'],[[label]+[fmt(frame[col].quantile(q),4) for q in [.05,.5,.95]] for col,label in [('h_cm','Hauteur (cm)'),('g_cm','Giron (cm)'),('blondel_cm','2h+g (cm)')]], [5*cm,4.2*cm,4.2*cm,4.2*cm])]
    story += [par('Méthode et limites','HeadStair'),par('Le modèle suppose une volée droite, n hauteurs identiques et n-1 girons identiques. L’angle représente la ligne entre les nez de marche : atan(h/g), et non l’angle de la diagonale de l’enveloppe atan(H/R). Les paliers, marches balancées, trémie, échappée, garde-corps, charges, épaisseur de paillasse et armatures restent à étudier séparément.'),
              par('H, R et W sont indépendants dans la simulation. La loi uniforme est symétrique autour de la référence ; la loi triangulaire est symétrique avec le mode à la référence. Les amplitudes sont des hypothèses ou tolérances documentées, pas des écarts-types ni des intervalles de confiance. Les relevés ne changent pas automatiquement ces lois. Aucun changement automatique de n au cours d’un tirage.'),
              par('Les critères sont ceux saisis par l’utilisateur. Ils ne constituent pas une validation réglementaire locale ni une approbation de construction. La probabilité affichée concerne ces critères géométriques, pas un risque d’accident ou de ruine.'),
              par('Notes : '+(p['notes'] or 'Aucune note complémentaire.')),par('Référence pédagogique : '+REFS[0][1])]
    def footer(canvas,doc):
        canvas.setFillColor(colors.HexColor(WINE));canvas.rect(0,29.7*cm-.25*cm,21*cm,.25*cm,stroke=0,fill=1)
        canvas.setFillColor(colors.HexColor(INK));canvas.setFont('StairVera',8);canvas.drawString(1.7*cm,1*cm,'EscalierLab · '+('FICTIF · ' if p['demo'] else '')+'Étude géométrique');canvas.drawRightString(19.3*cm,1*cm,str(doc.page))
    doc.build(story,onFirstPage=footer,onLaterPages=footer);return out.getvalue()


CSS='''<style>
.stApp{background:#FAF5ED;color:#302733}[data-testid="stMainBlockContainer"]{max-width:1480px;padding-top:4.8rem;padding-bottom:3rem}
[data-testid="stSidebar"]{background:#3E1D2B;color:#FAF1E8}[data-testid="stSidebar"] *{--text-color:#FAF1E8}[data-testid="stSidebarNav"] a span{color:#FAF1E8}[data-testid="stSidebarNav"] a:hover{background:#74374F}
[data-testid="stSidebar"] button{background:#F5E7DC;color:#4C2334;border-color:#D9B6A6}
h1,h2,h3{color:#402332;letter-spacing:-.03em}h1{font-size:2.3rem!important}p,label{line-height:1.65}
.mast{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;margin-bottom:25px}.brand{font-weight:900;font-size:1.25rem;letter-spacing:.08em;color:#542439}.brand span{color:#A95830}.tag{font-size:.67rem;letter-spacing:.12em;font-weight:750;background:#EAD5C7;color:#542439;border-radius:7px;padding:8px 12px}
.hero{display:grid;grid-template-columns:1.05fr 1fr;align-items:center;gap:20px;background:radial-gradient(at 100% 5%,#A15B65 0%,#702D46 48%,#492237 100%);border-radius:8px 32px 8px 32px;padding:35px 38px;color:#FFF6EC;box-shadow:0 16px 42px #6B2B4219;margin-bottom:26px}
.hero h2{font-family:Georgia,serif;font-weight:500;font-size:2.8rem;color:#FFF6EC;line-height:1.22;margin:13px 0 17px}.hero p{color:#F2DED5;max-width:510px}.eyebrow{font-size:.68rem;letter-spacing:.17em;font-weight:750;color:#ECC8AF}.hero img{width:100%;max-height:290px}.hero small{color:#F0D4BB}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:10px 0 25px}.kpi{padding:22px;background:#FFFDF9;border:1px solid #E7D9CE;border-radius:16px;min-width:0}.kpi:nth-child(2){background:#702D46;color:#FFF4E9;border-color:#702D46}.kpi:nth-child(3){background:#EED7C6}.kpi:nth-child(4){background:#DFE6DA}.kpi label{display:block;font-size:.77rem;line-height:1.45;opacity:.85}.kpi strong{display:block;font-size:1.85rem;font-weight:800;letter-spacing:-.04em;line-height:1.3;margin:10px 0;overflow-wrap:anywhere}.kpi small{font-size:.72rem;display:block;line-height:1.55;opacity:.87}
.note{background:#F0E1D7;border-left:4px solid #A95830;padding:16px 20px;border-radius:0 12px 12px 0;line-height:1.65;color:#552C39;margin:12px 0 24px}.sidebrand{font-family:Georgia,serif;font-size:1.9rem;color:#FFF2E4;margin:15px 0}.sidenote{font-size:.78rem;line-height:1.7;color:#E1BDAF;margin-bottom:22px}
.criterion{display:flex;justify-content:space-between;gap:14px;padding:12px 0;border-bottom:1px solid #E2D4C8;font-size:.86rem}.criterion b{white-space:normal}.ok{color:#356046}.bad{color:#8B2949}.unknown{color:#756B58}.foot{font-size:.75rem;color:#76655F;border-top:1px solid #DDCDC1;margin-top:32px;padding-top:23px}
[data-testid="stForm"]{background:#FFFDF9;border-color:#DECCBE;border-radius:17px;padding:22px}[data-testid="stButton"] button[kind="primary"],[data-testid="stFormSubmitButton"] button[kind="primary"]{background:#702D46;color:white;border-color:#702D46;border-radius:10px;min-height:44px}button p{white-space:normal;overflow-wrap:anywhere}a{overflow-wrap:anywhere}
@media(max-width:950px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.hero{padding:28px;grid-template-columns:1fr .85fr}.hero h2{font-size:2.25rem}}
@media(max-width:600px){[data-testid="stMainBlockContainer"]{padding:4.5rem 1rem 2rem}.hero{grid-template-columns:1fr;padding:25px}.hero h2{font-size:2.1rem}.hero img{max-height:190px}.kpi{padding:16px}.kpi strong{font-size:1.4rem}.kpis{gap:10px}h1{font-size:1.9rem!important}.criterion{font-size:.8rem}.tag{font-size:.6rem}.brand{font-size:1.07rem}}
</style>'''


def hero_art():
    parts=['<svg xmlns="http://www.w3.org/2000/svg" width="480" height="300" viewBox="0 0 480 300"><defs><pattern id="g" width="30" height="30" patternUnits="userSpaceOnUse"><path d="M30 0H0V30" fill="none" stroke="#F0CDB4" stroke-opacity=".12"/></pattern></defs><rect width="480" height="300" fill="url(#g)"/><path d="M50 260H430" stroke="#EDC5AC" stroke-width="1"/><ellipse cx="240" cy="260" rx="175" ry="22" fill="#31162B" opacity=".4"/>']
    for i in range(9):
        x=65+i*30;y=245-i*21
        parts.append(f'<path d="M{x} {y}l65 -32 30 9 -65 32Z" fill="#F3DCC2"/><path d="M{x} {y}l30 9v21l-30 -9Z" fill="#C18C74"/><path d="M{x+30} {y+9}l65 -32v21l-65 32Z" fill="#E3B79A"/>')
    parts.append('<path d="M58 220V56M53 61l5 -5 5 5M53 215l5 5 5 -5" fill="none" stroke="#EDC5AC"/><text x="35" y="145" fill="#FFE8CE" font-family="sans-serif" font-size="14">H</text><rect x="264" y="15" width="178" height="32" rx="5" fill="#421F32"/><text x="353" y="36" text-anchor="middle" fill="#F7DDC6" font-family="sans-serif" font-size="12">h · g · 2h + g</text></svg>')
    return 'data:image/svg+xml;base64,'+base64.b64encode(''.join(parts).encode()).decode()


def plot(fig,alt,key=None,three=False):
    fig.update_layout(template='plotly_white',font=dict(family='Arial',size=12,color=INK),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                      margin=dict(l=14,r=14,t=45,b=40),height=430 if three else 390,colorway=[WINE,COPPER,'#748A69'],legend=dict(orientation='h',y=1.02,yanchor='bottom',x=0))
    if three:fig.update_layout(margin=dict(l=10,r=10,t=15,b=15))
    st.plotly_chart(fig,width='stretch',alt=alt,key=key,config={'displaylogo':False,'scrollZoom':False})


def note(message):st.html('<div class="note">'+message+'</div>')


def kpis(items):
    st.html('<div class="kpis">'+''.join('<div class="kpi"><label>'+html.escape(str(label))+'</label><strong>'+html.escape(str(value))+'</strong><small>'+html.escape(str(detail))+'</small></div>' for label,value,detail in items)+'</div>')


def current():return st.session_state.demo_project if st.session_state.demo else st.session_state.real_project


def commit(p):
    p=validate(p)
    if st.session_state.demo:st.session_state.demo_project=p
    else:st.session_state.real_project=p
    st.session_state.rev+=1;st.session_state.sim=None;st.session_state.flash='Données enregistrées dans cette session.';st.rerun()


def actual_sim(p):
    s=st.session_state.get('sim');return s if s and s['fingerprint']==fingerprint(p) else None


def title(name,subtitle):st.title(name);st.caption(subtitle)


def dashboard():
    p=current();d=geometry(p['H'],p['R'],p['W'],p['n']);checks=diagnostic(p)
    st.html('<div class="hero"><div><div class="eyebrow">ATELIER DE CONCEPTION / VOLÉE DROITE</div><h2>Chaque marche<br>a sa mesure.</h2><p>Dimensionnez, comparez et comprenez les proportions de votre escalier, du premier relevé aux scénarios d’incertitude.</p><small>'+('DÉMONSTRATION ENTIÈREMENT FICTIVE' if p['demo'] else 'DONNÉES DU PROJET' if p['source'] else 'DIMENSIONS DU PROJET À DOCUMENTER')+'</small></div><img src="'+hero_art()+'" alt="Illustration axonométrique d’une volée droite et de sa hauteur"/></div>')
    kpis([('Hauteurs',p['n'],f"{p['n']-1} girons · sol supérieur d’arrivée"),('Hauteur de marche',fmt(d['h_cm'])+' cm','h = H / n'),('Giron',fmt(d['g_cm'])+' cm','g = R / (n − 1)'),('Emprise',fmt(d['footprint_m2'])+' m²','Volée seule · hors paliers')])
    if not p['source']:note('Les dimensions affichées sont des propositions à documenter dans « Dimensionner ». Les contrôles restent <b>à documenter</b>.')
    elif all(conditions(p).values()):note('La géométrie de référence respecte les <b>cinq critères saisis</b>. L’échappée, la trémie et l’étude structurelle restent à traiter pour le projet.')
    else:note('La géométrie présente <b>'+str(sum(not bool(v) for v in conditions(p).values()))+' écart(s)</b> aux critères saisis. Comparez d’autres nombres de hauteurs ou revoyez les dimensions disponibles.')
    a,b=st.columns([1.4,1])
    with a:
        st.subheader('La volée en profil');plot(profile_figure(p),'Profil coté de la volée, hauteur totale et ligne des nez de marche',key='dashboard_profile')
        st.caption('La pente de foulée suit les nez de marche. Les girons s’arrêtent au plancher supérieur ; les paliers ne sont pas inclus dans R.')
    with b:
        st.subheader('Lecture des critères')
        for row in checks.itertuples():
            style='unknown' if row.etat=='À documenter' else 'ok' if row.etat=='Dans le critère' else 'bad'
            st.html('<div class="criterion"><span>'+html.escape(row.controle)+'<br><small>'+html.escape(row.critere)+'</small></span><b class="'+style+'">'+html.escape(row.valeur)+'<br><small>'+html.escape(row.etat)+'</small></b></div>')
        st.caption('Les seuils initiaux sont des critères d’exercice modifiables. Le seul respect de 2h + g ne suffit pas à qualifier tout l’escalier.')
    with st.container(horizontal=True):
        st.page_link(st.session_state.pages['dimensionner'],label='Ajuster les dimensions',icon=':material/architecture:')
        st.page_link(st.session_state.pages['comparer'],label='Comparer les configurations',icon=':material/compare_arrows:')


def dimension_page():
    p=current();rev=st.session_state.rev
    title('Dimensionner la volée','Une hauteur entre sols finis, un reculement disponible et un nombre entier de hauteurs.')
    with st.form('dimensions_'+str(rev)):
        name=st.text_input('Nom du projet',p['name'],max_chars=120)
        a,b=st.columns(2)
        H=a.number_input('Hauteur entre sols finis H (m)',.30,12.,p['H'],step=.01,format='%.3f')
        R=b.number_input('Reculement de la volée R (m ; hors paliers)',.30,20.,p['R'],step=.01,format='%.3f')
        W=a.number_input('Largeur utile W (m)',.30,6.,p['W'],step=.01,format='%.3f')
        n=b.number_input('Nombre de hauteurs n',2,60,p['n'])
        kind=st.selectbox('Nature des dimensions',KINDS,index=KINDS.index(p['kind']))
        source=st.text_area('Source des dimensions ou justification de l’hypothèse',p['source'],max_chars=800)
        notes=st.text_area('Notes du projet',p['notes'],max_chars=1200)
        st.caption('Une volée droite : n hauteurs uniformes, n−1 girons uniformes, le sol fini supérieur servant d’arrivée. Enregistrer de nouvelles dimensions remet les lois d’incertitude sur Fixe pour éviter des bornes devenues inadaptées.')
        if st.form_submit_button('Calculer et enregistrer',type='primary',icon=':material/calculate:'):
            q=copy.deepcopy(p);q.update(name=name,H=H,R=R,W=W,n=n,kind=kind,source=source,notes=notes)
            if any(q[k]!=p[k] for k in ['H','R','W']):q['uncertainty']=fixed_uncertainty()
            commit(q)
    d=geometry(p['H'],p['R'],p['W'],p['n'])
    kpis([('Hauteur h',fmt(d['h_cm'])+' cm','100 × H / n'),('Giron g',fmt(d['g_cm'])+' cm','100 × R / (n − 1)'),('Pas 2h + g',fmt(d['blondel_cm'])+' cm','Relation de proportions'),('Pente de foulée',fmt(d['angle_deg'])+'°','atan(h / g)')])
    with st.expander('Modifier les critères du projet'):
        c=p['criteria']
        with st.form('criteria_'+str(rev)):
            values={};cols=st.columns(2)
            for i,(key,label) in enumerate([('h_min','Hauteur minimale (cm)'),('h_max','Hauteur maximale (cm)'),('g_min','Giron minimal (cm)'),('g_max','Giron maximal (cm)'),('b_min','2h + g minimal (cm)'),('b_max','2h + g maximal (cm)')]):
                values[key]=cols[i%2].number_input(label,.1,200.,c[key],step=.5)
            values['angle_max']=cols[0].number_input('Pente de foulée maximale (degrés)',1.,89.,c['angle_max'])
            values['width_min']=cols[1].number_input('Largeur utile minimale (m)',.1,6.,c['width_min'],step=.05)
            source=st.text_area('Référence et contexte des critères',p['criteria_source'],max_chars=800)
            st.caption('Ces critères d’exercice ne sont pas un ensemble de règles réglementaires locales. Documenter les exigences applicables au bâtiment et à son usage.')
            if st.form_submit_button('Enregistrer les critères'):
                q=copy.deepcopy(p);q.update(criteria=values,criteria_source=source);commit(q)
    with st.expander('Contrôler une série de hauteurs relevées'):
        m=p['measures'];st.write('Les relevés sont des mesures individuelles en cm. Ils ne remplacent pas automatiquement la géométrie ni les lois probabilistes.')
        with st.form('measures_'+str(rev)):
            source=st.text_input('Source des relevés',m['source'],max_chars=800)
            raw=st.text_area('Hauteurs relevées en cm : une par ligne ou séparées par ;','; '.join(str(x).replace('.',',') for x in m['values']))
            a,b=st.columns(2)
            spread=a.number_input('Écart min–max toléré pour cet exercice (cm)',.01,20.,m['spread_cm'])
            closure=b.number_input('Écart de fermeture toléré pour cet exercice (cm)',.001,20.,m['closure_cm'],format='%.3f')
            if st.form_submit_button('Analyser et conserver les relevés'):
                vals=[num(x.strip(),'Hauteur relevée',.01,100.) for x in re.split('[;\n]',raw) if x.strip()]
                q=copy.deepcopy(p);q['measures']=dict(source=source,values=vals,spread_cm=spread,closure_cm=closure);commit(q)
        ms=measure_stats(p)
        if ms:
            kpis([('Mesures',ms['count'],f"{p['n']} attendues pour une fermeture complète"),('Moyenne',fmt(ms['mean'],3)+' cm','Relevés observés'),('Écart-type',fmt(ms['std'],3)+' cm' if ms['std'] is not None else 'Non défini','Diviseur : nombre de relevés − 1'),('Écart min–max',fmt(ms['spread'],3)+' cm','Étendue des hauteurs relevées')])
            st.write('**Fermeture :** '+(fmt(ms['closure'],3)+' cm, '+('dans la tolérance d’exercice.' if ms['closure_ok'] else 'hors tolérance d’exercice.') if ms['complete'] else 'non évaluée : série incomplète ou nombre de mesures différent de n.'))
            st.write('**Régularité de la série :** '+('dans le critère d’étendue saisi.' if ms['spread_ok'] else 'étendue supérieure au critère saisi.'))
            fig=go.Figure(go.Scatter(x=list(range(1,ms['count']+1)),y=m['values'],mode='lines+markers',name='Relevés'))
            fig.add_hline(y=100*p['H']/p['n'],line_dash='dot',line_color=COPPER);fig.update_xaxes(title='Numéro de hauteur');fig.update_yaxes(title='Hauteur relevée · cm');plot(fig,'Hauteurs individuelles relevées et hauteur uniforme de référence')
            st.caption('La fermeture somme les hauteurs relevées et soustrait H uniquement si leur nombre correspond à n. Une petite série ne caractérise pas tous les écarts d’exécution.')
        st.subheader('Importer les relevés')
        st.download_button('Modèle CSV fictif',csv_bytes(pd.DataFrame(dict(numero=range(1,19),hauteur_cm=[16.67]*18))),'EscalierLab_modele_fictif.csv','text/csv',on_click='ignore')
        upload=st.file_uploader('Relevés CSV ou Excel',type=['csv','xlsx'],key='upload_measure_'+str(rev))
        source=st.text_input('Source du fichier de relevés',max_chars=800,key='import_source_'+str(rev))
        if upload:
            try:
                vals=import_measures(upload.getvalue(),upload.name)
                st.dataframe(pd.DataFrame(dict(numero=range(1,len(vals)+1),hauteur_cm=vals)),hide_index=True,alt='Aperçu des hauteurs importées')
                sure=st.checkbox('Remplacer les relevés actuels par ce fichier')
                if st.button('Importer les relevés vérifiés',disabled=not sure):
                    q=copy.deepcopy(p);q['measures'].update(values=vals,source=source);commit(q)
            except Exception as ex:st.error('Import non appliqué : '+str(ex))


def comparison_page():
    p=current();title('Comparer les configurations','Même hauteur, même reculement et même largeur : seul le nombre de hauteurs change.')
    a,b=st.columns(2);lo=a.number_input('Minimum de hauteurs à explorer',2,60,10);hi=b.number_input('Maximum de hauteurs à explorer',2,60,30)
    if lo>hi:st.info('Le minimum doit être inférieur ou égal au maximum.');return
    df=candidates(p,lo,hi);good=df[df.respect_geometrique].sort_values('distance_pas_cible_cm')
    if not p['source']:note('Comparaison de dimensions <b>non documentées</b>. Les compatibilités calculées sont exploratoires jusqu’à déclaration de leur source.')
    kpis([('Configurations',len(df),'Nombres entiers de hauteurs'),('Dans les critères',len(good),'Les cinq contrôles géométriques'),('Pas cible',fmt((p['criteria']['b_min']+p['criteria']['b_max'])/2)+' cm','Milieu des bornes 2h + g'),('Choix actuel',p['n'],'Hauteurs conservées au projet')])
    fig=go.Figure(go.Scatter(x=df.giron_cm,y=df.hauteur_cm,mode='markers+text' if len(df)<=25 else 'markers',text=df.hauteurs.astype(str),textposition='top center',marker=dict(size=11,color=[WINE if ok else '#B6AAA1' for ok in df.respect_geometrique]),hovertemplate='Giron %{x:.2f} cm<br>Hauteur %{y:.2f} cm<extra></extra>'))
    c=p['criteria'];fig.add_shape(type='rect',x0=c['g_min'],x1=c['g_max'],y0=c['h_min'],y1=c['h_max'],fillcolor='#DDE6D2',opacity=.35,line_width=0,layer='below')
    fig.update_xaxes(title='Giron · cm');fig.update_yaxes(title='Hauteur de marche · cm');plot(fig,'Configurations par nombre de hauteurs ; zone de bornes h et g, points bordeaux satisfaisant les cinq critères')
    st.caption('La zone colorée montre seulement les bornes h et g. Le statut de chaque point tient aussi compte de 2h+g, de la pente et de la largeur. Les étiquettes indiquent n.')
    st.dataframe(df,hide_index=True,alt='Comparaison des configurations et écarts à la cible de pas')
    selection=st.selectbox('Nombre de hauteurs à adopter',df.hauteurs.tolist(),index=df.hauteurs.tolist().index(p['n']) if p['n'] in df.hauteurs.values else 0)
    row=df[df.hauteurs==selection].iloc[0]
    st.write(f"Choix : **{selection} hauteurs**, h = **{fmt(row.hauteur_cm)} cm**, g = **{fmt(row.giron_cm)} cm**.")
    if not row.respect_geometrique:st.warning('Cette configuration présente un écart aux critères. Vous pouvez la conserver comme scénario à étudier.')
    if st.button('Adopter cette configuration',type='primary'):
        q=copy.deepcopy(p);q['n']=int(selection);commit(q)
    st.download_button('Tableau comparatif CSV',csv_bytes(df),'EscalierLab_comparaison.csv','text/csv',on_click='ignore')
    st.caption('Le classement PDF des solutions dans les critères utilise la proximité de 2h+g au pas cible. Aucun classement ne valide la structure ni l’usage réel.')


def drawings_page():
    p=current();title('Schémas et proportions','Profil, plan et volume géométrique construits à partir des dimensions enregistrées.')
    view=st.segmented_control('Vue',['Profil','Plan','Volume 3D'],default='Profil')
    if view=='Plan':plot(plan_figure(p),'Plan de la volée, girons, largeur et sens de montée')
    elif view=='Volume 3D':
        plot(volume_figure(p),'Volume géométrique 3D des marches et contremarches, sans épaisseur structurelle',three=True)
        st.caption(f"Dimensions en mètres : H = {fmt(p['H'])}, R = {fmt(p['R'])}, W = {fmt(p['W'])}. Glisser pour tourner le volume ; utiliser les commandes du graphique pour zoomer. La représentation ne décrit ni matériau, ni paillasse, ni armatures.")
    else:plot(profile_figure(p),'Profil de la volée à marches uniformes et ligne de foulée')
    d=geometry(p['H'],p['R'],p['W'],p['n'])
    note(f"<b>{p['n']} hauteurs</b> de {fmt(d['h_cm'])} cm ; <b>{p['n']-1} girons</b> de {fmt(d['g_cm'])} cm. Le plancher d’arrivée tient lieu de dernière marche. Emprise : {fmt(d['footprint_m2'])} m², hors paliers.")
    st.caption('Les axes du profil et du plan utilisent une échelle métrique commune. Le volume respecte les rapports géométriques. Une vue très étroite peut réduire la lisibilité des annotations ; les cotes restent dans les indicateurs et le PDF.')


def uncertainty_page():
    p=current();rev=st.session_state.rev
    title('Incertitudes et marges','Quelle fréquence de scénarios respecte les critères quand les dimensions varient ?')
    if not p['source']:st.info('Documentez les dimensions dans « Dimensionner » avant de lancer les simulations.');return
    note(f"Le nombre de hauteurs reste fixé à <b>{p['n']}</b> pour tous les scénarios. H, R et W sont supposés indépendants. On étudie les critères géométriques saisis ; les variations ne sont pas une évolution de l’escalier dans le temps.")
    with st.form('uncertainty_'+str(rev)):
        settings={}
        for key,label in [('H','Hauteur entre sols finis'),('R','Reculement disponible'),('W','Largeur utile')]:
            st.markdown(f"**{key} · {label} : référence {fmt(p[key],3)} m**")
            a,b=st.columns(2);old=p['uncertainty'][key]
            law=a.selectbox('Loi · '+key,LAWS,index=LAWS.index(old['law']))
            delta=b.number_input('Amplitude ± · '+key+' (cm)',0.,1000.,old['delta_cm'],step=.1,format='%.3f')
            settings[key]=dict(law=law,delta_cm=delta if law!='Fixe' else 0.)
        st.caption('Fixe : amplitude ignorée. Uniforme : bornes référence ± amplitude. Triangulaire symétrique : mêmes bornes, mode à la référence. Les amplitudes sont des hypothèses ou tolérances, pas des écarts-types.')
        if st.form_submit_button('Enregistrer les hypothèses',type='primary'):
            q=copy.deepcopy(p);q['uncertainty']=settings;commit(q)
    with st.form('simulate'):
        a,b=st.columns(2);count=a.selectbox('Nombre de tirages',[1000,10000,50000,100000],index=1);seed=b.number_input('Graine aléatoire',0,2**32-1,42)
        if st.form_submit_button('Lancer les simulations',type='primary',icon=':material/play_arrow:'):
            with st.spinner('Calcul des géométries possibles…'):frame=simulate(json.dumps(p,ensure_ascii=False,sort_keys=True),count,seed)
            st.session_state.sim=dict(df=frame,count=count,seed=seed,fingerprint=fingerprint(p))
    sim=actual_sim(p)
    if not sim:st.caption('Aucune simulation actuelle. Enregistrez les hypothèses, puis lancez le calcul.');return
    df=sim['df'];success=int(df.tous_criteres.sum());prob=success/len(df);lo,hi=wilson(success,len(df))
    kpis([('Tous les critères',fmt(100*prob,2)+' %',f"{success} / {len(df)} scénarios"),('Au moins un écart',fmt(100*(1-prob),2)+' %','Complément du résultat global'),('Giron médian',fmt(df.g_cm.median())+' cm','Distribution simulée'),('Wilson à 95 %',fmt(lo*100,3)+'–'+fmt(hi*100,3)+' %','Précision numérique de la fréquence')])
    st.caption(f"Résultat : {sim['count']} tirages, graine {sim['seed']}, n={p['n']}. Relancer pour appliquer d’autres réglages. Wilson ne mesure pas la validité des hypothèses.")
    a,b=st.columns(2)
    with a:
        fig=go.Figure(go.Histogram(x=df.g_cm,nbinsx=45,histnorm='probability',marker_color=WINE,name='Giron'))
        fig.add_vline(x=p['criteria']['g_min'],line_color=COPPER,line_dash='dash');fig.add_vline(x=p['criteria']['g_max'],line_color=COPPER,line_dash='dash');fig.update_xaxes(title='Giron simulé · cm');fig.update_yaxes(title='Fréquence');plot(fig,'Distribution simulée du giron et bornes du critère')
    with b:
        flags=['hauteur_ok','giron_ok','blondel_ok','pente_ok','largeur_ok'];labels=['Hauteur','Giron','2h + g','Pente','Largeur']
        fig=go.Figure(go.Bar(x=[100*(~df[k]).mean() for k in flags],y=labels,orientation='h',marker_color=COPPER));fig.update_xaxes(title='Scénarios hors critère · %',range=[0,100]);plot(fig,'Fréquences individuelles d’écart, non additives car plusieurs écarts peuvent coexister')
    stats=pd.DataFrame([dict(grandeur=label,moyenne=df[col].mean(),ecart_type=df[col].std(ddof=1),P05=df[col].quantile(.05),mediane=df[col].median(),P95=df[col].quantile(.95)) for col,label in [('h_cm','Hauteur (cm)'),('g_cm','Giron (cm)'),('blondel_cm','2h+g (cm)'),('pente_deg','Pente (degrés)'),('emprise_m2','Emprise (m²)')]])
    st.dataframe(stats,hide_index=True,alt='Statistiques et quantiles des géométries simulées')
    st.caption('P05–P95 décrit environ 90 % des résultats sous le modèle, pas un intervalle de confiance de la moyenne. Les fréquences d’écart par critère ne s’additionnent pas.')
    with st.container(horizontal=True):
        st.download_button('Tous les scénarios CSV',csv_bytes(df),'EscalierLab_simulations.csv','text/csv',on_click='ignore')
        data=dict(app='EscalierLab · résultats',project=p,count=sim['count'],seed=sim['seed'],probability=prob,wilson95=[lo,hi])
        st.download_button('Hypothèses de simulation JSON',json.dumps(data,ensure_ascii=False,indent=2).encode(),'EscalierLab_hypotheses.json','application/json',on_click='ignore')


def reports_page():
    p=current();s=actual_sim(p);title('Rapports et tests','Exporter une note de calcul et vérifier le fonctionnement avec des exemples connus.')
    st.caption('Les projets restent dans la session du navigateur. Téléchargez le JSON pour les conserver avant fermeture ou reconnexion. L’espace fictif est séparé de votre étude personnelle.')
    prefix='FICTIF_' if p['demo'] else ''
    with st.container(horizontal=True):
        st.download_button('Note de calcul PDF',pdf_bytes(p,s),prefix+'EscalierLab_note.pdf','application/pdf',type='primary',icon=':material/picture_as_pdf:',on_click='ignore')
        st.download_button('Classeur Excel',excel_bytes(p,s),prefix+'EscalierLab.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',on_click='ignore')
        st.download_button('Sauvegarde complète JSON',backup_bytes(p),prefix+'EscalierLab_projet.json','application/json',on_click='ignore')
        st.download_button('Relevés CSV',csv_bytes(measures_frame(p)),prefix+'EscalierLab_releves.csv','text/csv',on_click='ignore')
    st.caption('Excel : sept feuilles. PDF et Excel incluent une synthèse de la simulation actuelle, si elle existe. Les tirages complets s’exportent dans « Incertitudes ». La sauvegarde conserve les entrées et les relevés ; la simulation se recalcule avec la graine souhaitée.')
    with st.expander('Restaurer une étude'):
        file=st.file_uploader('Sauvegarde EscalierLab JSON',type=['json'])
        if file:
            try:
                restored=restore(file.getvalue());st.write(restored['name']+' · '+('EXEMPLE FICTIF' if restored['demo'] else 'ÉTUDE PERSONNELLE'))
                sure=st.checkbox('Remplacer le projet de cet espace par la sauvegarde')
                if st.button('Restaurer le projet',disabled=not sure):
                    st.session_state.demo=restored['demo'];commit(restored)
            except Exception as ex:st.error('Restauration non appliquée : '+str(ex))
    st.subheader('Comprendre le calcul')
    st.latex(r'h=\frac{H}{n},\quad g=\frac{R}{n-1},\quad B=2h+g,\quad\theta=\arctan\left(\frac{h}{g}\right)')
    st.write('H et R en mètres donnent h et g en mètres ; l’interface les convertit en cm. Une volée comporte n hauteurs et n−1 girons, le plancher d’arrivée tenant lieu de dernière marche. La pente de foulée suit les nez de marche : elle diffère de l’angle de la diagonale H/R de l’enveloppe.')
    st.subheader('Tests guidés')
    with st.container(horizontal=True):
        if st.button('Réinitialiser l’exemple fictif',type='primary'):
            st.session_state.demo=True;commit(new_project(True))
        if st.button('Ouvrir un exemple compact'):
            q=new_project(True);q.update(name='Volée compacte · exemple fictif',H=2.7,R=3.6,n=15);q['criteria']['g_min']=24.;q['uncertainty']=fixed_uncertainty();q['source']='Exemple compact entièrement fictif : H=2,70 m, R=3,60 m, n=15.';st.session_state.demo=True;commit(q)
    tests=[
        ('1 · Géométrie de référence','H=3 m ; R=4,76 m ; W=1 m ; n=18. Résultats : 17 girons ; h=16,6667 cm ; g=28 cm ; 2h+g=61,3333 cm ; pente de foulée ≈30,7627° ; emprise=4,76 m².'),
        ('2 · Comparaison','Avec les critères initiaux, explorer n=10 à 30 : seule la configuration n=18 satisfait les cinq critères. n=17 donne 2h+g≈65,0441 cm, au-dessus de 64 ; n=19 donne g≈26,4444 cm, sous le minimum 28.'),
        ('3 · Probabilité connue','Exemple de référence : H=3 m et W=1 m fixes ; R uniforme entre 4,72 et 4,80 m (amplitude ±4 cm) ; n=18 fixe. Le giron est ≥28 cm si R≥4,76 m, soit 50 % exactement sous ces hypothèses. Les autres critères restent satisfaits sur cet intervalle. À 10000 tirages, environ ±0,98 point autour de 50 % dans 95 % des répétitions, sans garantie pour chaque graine.'),
        ('4 · Tout fixe','Dans Incertitudes, mettre H, R et W sur Fixe, enregistrer, puis simuler : 100 % des scénarios dans les critères pour n=18. Passer à n=17 via Comparer et relancer : 0 % pour le résultat global avec ces mêmes dimensions.'),
        ('5 · Relevés observés','Importer le modèle fictif : 18 hauteurs de 16,67 cm. Moyenne=16,67 cm ; écart-type=0 ; étendue=0 ; somme=300,06 cm ; fermeture=+0,06 cm. Ces relevés ne remplacent pas les dimensions ni les lois de simulation.'),
        ('6 · Série incomplète','Saisir 16 ;16,5 ;17 ;17,5, avec source fictive. Moyenne=16,75 cm ; écart-type échantillonnal≈0,645497 cm ; étendue=1,5 cm. La fermeture reste non évaluée car quatre valeurs ne décrivent pas les 18 hauteurs.'),
        ('7 · Erreurs et sauvegarde','Un minimum de critère supérieur au maximum doit être refusé. Une amplitude conduisant à une dimension nulle ou négative est refusée. Télécharger puis restaurer le JSON doit retrouver dimensions, critères, hypothèses et relevés sans modifier l’autre espace.')]
    for label,body in tests:
        with st.expander(label):st.write(body)
    st.subheader('Périmètre de cette première version')
    st.write('Une volée droite et uniforme. Les paliers, marches tournantes, trémie, échappée, garde-corps et éléments porteurs ne sont pas dimensionnés. L’épaisseur, les armatures et les charges exigent une étude structurelle distincte. Les critères initiaux sont des hypothèses pédagogiques, et le logiciel ne déduit pas une conformité locale.')
    for label,url in REFS:st.markdown(f'[{label}]({url})')


def main():
    st.set_page_config(page_title='EscalierLab · Chaque marche a sa mesure',page_icon=':material/stairs:',layout='wide');st.html(CSS)
    ss=st.session_state
    if 'real_project' not in ss:ss.update(real_project=new_project(),demo_project=new_project(True),demo=True,rev=0,sim=None,flash='')
    p=current();st.html('<div class="mast"><div class="brand">ESCALIER<span>LAB</span></div><span class="tag">'+('EXEMPLE FICTIF' if ss.demo else 'ÉTUDE PERSONNELLE')+' / GÉOMÉTRIE</span></div>')
    routes=[('accueil',dashboard,'Tableau de bord','dashboard'),('dimensionner',dimension_page,'Dimensionner','architecture'),('comparer',comparison_page,'Comparer','compare_arrows'),('schemas',drawings_page,'Schémas','view_in_ar'),('incertitudes',uncertainty_page,'Incertitudes','query_stats'),('rapports',reports_page,'Rapports et tests','description')]
    pages={path:st.Page(func,title=label,icon=':material/'+icon+':',url_path=path,default=path=='accueil') for path,func,label,icon in routes};ss.pages=pages
    with st.sidebar:st.html('<div class="sidebrand">EscalierLab.</div><div class="sidenote">La géométrie en mouvement.<br>Le calcul à portée de main.</div>')
    nav=st.navigation(list(pages.values()))
    with st.sidebar:
        st.caption('ATELIER ACTIF');st.write('**'+p['name']+'**')
        if ss.demo:
            if st.button('Revenir à mon étude',icon=':material/folder:'):ss.demo=False;ss.sim=None;ss.rev+=1;st.rerun()
        else:
            if st.button('Ouvrir l’exemple fictif',icon=':material/science:'):ss.demo=True;ss.sim=None;ss.rev+=1;st.rerun()
        st.caption('Volée droite · Version 1.0 · Sans compte')
    if ss.flash:st.toast(ss.flash,icon=':material/check_circle:');ss.flash=''
    try:nav.run()
    except (ValueError,TypeError,KeyError,OverflowError) as ex:st.error('Action non appliquée : '+str(ex))
    st.html('<div class="foot">EscalierLab · Dimensionnement géométrique pédagogique selon les critères saisis. Les dispositions constructives et l’étude structurelle restent à examiner.</div>')


if __name__=='__main__':main()
