"""Build an editable, source-reconciled Signal Desk business plan. Requires python-pptx."""
from pathlib import Path
import re, json, math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN

ROOT=Path(__file__).resolve().parents[1]
OUT=Path('/tmp/business-plan'); OUT.mkdir(exist_ok=True)
P=Presentation(); P.slide_width=Inches(13.333); P.slide_height=Inches(7.5)
C={'ink':'162B3A','blue':'0068A0','green':'008A68','white':'FFFFFF','muted':'62717D','light':'EFF4F7','line':'D6E0E5','red':'B74444'}
slides=[]
def text(s,t,x,y,w,h,size=22,bold=False,color='ink'):
 box=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=box.text_frame; tf.word_wrap=True
 tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
 for i,line in enumerate(str(t).split('\n')):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line;p.font.name='Arial';p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=RGBColor.from_string(C.get(color,color));p.space_after=Pt(8)
 return box

def rect(s,x,y,w,h,color):
 sh=s.shapes.add_shape(1,Inches(x),Inches(y),Inches(w),Inches(h));sh.fill.solid();sh.fill.fore_color.rgb=RGBColor.from_string(C[color]);sh.line.fill.background();return sh

def new(title,section,sub='',dark=False):
 s=P.slides.add_slide(P.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string(C['ink' if dark else 'white'])
 fg='white' if dark else 'ink';text(s,'SIGNAL DESK  /  '+section.upper(),.55,.35,11,.3,12,True,'green' if dark else 'blue')
 text(s,title,.55,.95,12.1,1.25,36,True,fg)
 if sub:text(s,sub,.55,2.15,12.1,.65,17,False,'light' if dark else 'muted')
 text(s,'BUSINESS PLAN • OCTOBER 2026  |  CONFIDENTIAL • MANAGEMENT PLANNING, NOT AUDITED',.55,7.03,11,.25,9,False,'light' if dark else 'muted')
 text(s,f'{len(P.slides):02d}',12.1,7.01,.65,.3,11,False,fg)
 slides.append({'title':title,'section':section,'summary':sub});return s

def blocks(title,section,items,sub=''):
 s=new(title,section,sub); n=len(items);w=12.1/n
 for i,(head,body) in enumerate(items):
  x=.55+i*w;text(s,f'{i+1:02}',x,3,.6,.5,22,True,'green');text(s,head,x,3.65,w-.4,.7,24,True);text(s,body,x,4.45,w-.4,2.05,20)
 return s

def table(title,section,headers,rows,sub='',size=18):
 s=new(title,section,sub);n=len(rows)+1;shape=s.shapes.add_table(n,len(headers),Inches(.55),Inches(3),Inches(12.1),Inches(3.6));tb=shape.table
 tb.columns[0].width=Inches(4.1)
 for col in list(tb.columns)[1:]:col.width=Inches(8/(len(headers)-1))
 for r,row in enumerate([headers]+rows):
  for j,value in enumerate(row):
   cell=tb.cell(r,j);cell.text=str(value);cell.margin_left=Inches(.12);cell.margin_top=Inches(.065);cell.margin_bottom=Inches(.035)
   cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string(C['ink' if r==0 else ('light' if r%2 else 'white')])
   for p in cell.text_frame.paragraphs:
    p.font.name='Arial';p.font.size=Pt(size);p.font.bold=r==0;p.font.color.rgb=RGBColor.from_string(C['white' if r==0 else 'ink'])
 return s

def chart(s,series,x=.65,y=3,w=8,h=3.55):
 d=CategoryChartData();d.categories=['Y1','Y2','Y3','Y4','Y5']
 for name,values in series:d.add_series(name,values)
 ch=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Inches(x),Inches(y),Inches(w),Inches(h),d).chart
 ch.has_legend=True;ch.legend.position=XL_LEGEND_POSITION.BOTTOM;ch.legend.font.name='Arial';ch.legend.font.size=Pt(12)
 for i,ser in enumerate(ch.series):ser.format.fill.solid();ser.format.fill.fore_color.rgb=RGBColor.from_string(C[['blue','green','red'][i%3]])
 ch.category_axis.tick_labels.font.name='Arial';ch.category_axis.tick_labels.font.size=Pt(12);ch.value_axis.tick_labels.font.size=Pt(12)
 return ch

rev=[7700,41000,117500,287000,581000];coun=[3000,12000,33000,78000,150000];biz=[1500,6000,16500,39000,81000];mun=[0,5000,20000,50000,110000];data=[2000,12000,30000,75000,150000];ads=[1200,6000,18000,45000,90000]
cogs=[2600,5500,11600,23000,40000];salary=[225000,270000,315000,405000,450000];rent=[9000,10800,12600,16200,18000];hw=[7500,4000,4500,6500,6000];tools=[6000,7200,8400,10800,12000];gtm=[6000,12000,20000,35000,50000]
opex=[sum(x) for x in zip(salary,rent,hw,tools,gtm)];ebitda=[r-c-o for r,c,o in zip(rev,cogs,opex)];free=[r-c for r,c in zip(rev,coun)];free_ebitda=[r-c-o for r,c,o in zip(free,cogs,opex)]
def money(n):return f'({abs(n)/1000:,.1f})' if n<0 else f'{n/1000:,.1f}'
def row(label,a):return [label]+[money(x) for x in a]
H=['US$000 unless stated','Y1','Y2','Y3','Y4','Y5']

s=new('SIGNAL DESK','Strategic business plan','Community intelligence. Municipal action.',True)
photo=Path('/tmp/business-plan/city-000.jpg')
if photo.exists():
 s.shapes.add_picture(str(photo),Inches(7.8),Inches(2.9),width=Inches(4.95),height=Inches(3.45))
text(s,'Strategic growth &\ncapital investment plan',.55,3.15,7,1.5,32,True,'white');text(s,'Five-year outlook • South Africa to regional scale\nPrepared from three management-supplied documents',.55,5.15,6.7,1,18,False,'light')
blocks('A validated intake foundation; a commercial model to prove','Executive summary', [('Evidence','Six-year Ward 115 history supports the data proposition. Recorded reports and founder-reported outcomes are not audited repairs.'),('Strategy','Replicate up to five wards; sell sponsor-funded and institutional analysis while keeping core civic access free.'),('Capital','The supplied paid-seat forecast needs about $981k before working capital. A $100k raise is a first tranche, not full scale funding.')])
blocks('Read the plan through three decision lenses','Contents',[('01 / Business','Evidence, workflow, market, offers, competition and first-cohort acquisition.'),('02 / Economics','Source forecast, free-core reconciliation, cash flow, balance sheet, monthly bridge and sensitivities.'),('03 / Execution','Funding gates, team, compliance, expansion and the complete supplied text appendix.')])
table('Conflicting source assumptions are disclosed, not blended','Source reconciliation',['Topic','Text / pitch deck','Expansion plan','Treatment'],[['Councillor pricing','$10 per month','Core service free','Free-core strategic case; paid-seat reference'],['Market size','$47.4m / $1.11tn TAM','Coverage counts, not revenue','Bottom-up buyer validation'],['Runway','$100k; text says 12–14 months','Funded scopes and gates','4.8-month reference burn'],['Resolution claims','100 resolved; <30-min admin','Independent feedback sparse','Founder-reported, not independently audited']],size=16,sub='S1: supplied text. S2: pitch deck, 15 Oct 2026. S3: supplied growth and expansion plan.')
blocks('Fragmented signals create a handoff problem','Problem',[('Capture gap','Chat messages, locations and evidence sit across disconnected community channels.'),('Coordination burden','The pitch reports up to four hours of daily manual work per ward office. A before/after time study is still required.'),('Planning blind spot','Repeated failures cannot reliably inform budgets until reports, incidents and official cases are linked.')])
table('Ward 115 provides depth, not proof of paid demand','Evidence',['Historical measure','Source value','Interpretation'],[['All messages','155,284','Source volume over six years'],['Service reports','30,977','Reports before incident consolidation'],['Councillor messages','50,899','Coordination; subset, not additive'],['Unique community members','5,563','Reported people, not households'],['Groups / coverage','5 / 6 years','Historical archive scope']],sub='S3 pp. 3–4; historical units must not be added together.',size=19)
table('Pilot periods and current intake must stay distinct','Evidence',['Measure','Value','Evidence status'],[['June 30–July 31 users / reports','6,500 / 577','Founder-reported, S2 p. 6'],['Ward-level resolved / AI spend','100 / ~$11','Independent outcome audit required'],['Photo-video / matched municipal IDs','61 / 139','Period-specific pilot'],['Current non-archived reports','3,048','S3 review dated 15 Oct 2026'],['Real refs / coordinates','848 / 1,246','27.8% / 40.9% of current reports']],sub='Do not reconcile different dates or definitions into one traction claim.',size=18)
blocks('Capture → link → act → verify → learn','Product architecture',[('Capture & link','Telegram-first and web intake. Preserve original messages, consent, evidence and location; link duplicates to canonical incidents.'),('Act & verify','Assign service ownership and real municipal references. Record contractor evidence and independent feedback; reopen disputed cases.'),('Learn','Separate reports, incidents, municipal cases and rewards. Build governed aggregate evidence for planning and resilience.')])
blocks('Expand the orchestration layer, not municipal infrastructure','Technology',[('Live foundation','Five Telegram groups, Ward 115 mini app, AI triage and case history; WhatsApp administrative console. Architecture is founder-supplied.'),('Next release','Multi-tenant isolation, cross-channel linking, routing validation, accessible capture and reusable location records.'),('Control points','Human review for low confidence, access controls, audit history and disaster recovery. SMS/USSD remains a proposed capability.')])
table('Sell evidence to fund free civic participation','Offer portfolio',['Offer','Buyer','Commercial validation'],[['Resident & councillor tools','Residents / councillors','Free core access; S3'],['Ward intelligence baseline','Sponsor / institution','Paid bounded scope and acceptance criteria'],['Planning evidence pack','Sponsor / municipality','Recurring analytical service; price unvalidated'],['Municipal intelligence','Authorised municipal team','$10k/year reference assumption from S1'],['Aggregate insight / API','Utilities / planners / AI buyers','Lawful rights and paid data agreement']],sub='No signed commercial agreements or validated sponsor prices supplied.',size=18)
blocks('A grounded market starts with identifiable budget owners','Market',[('South Africa','4,488 wards, 9,473 councillors and 257 municipalities have different dates and units. They are coverage, not paying accounts.'),('Reference capacity','$1.137m councillor + $2.570m municipal annual revenue at S1 prices; $3.707m subtotal before businesses or data. Not validated TAM.'),('Demand test','Name the buyer, procurement route, budget, use case and payment terms. Validate contract value before extrapolating a regional market.')])
table('Keep market estimates separate from obtainable revenue','Market',['Source estimate','Supplied figure','Assessment'],[['S1 continental TAM','$47.4m','Method and ward-level prices not provided'],['S1 South African SAM','$6.1m','Includes undefined adjacent buyers'],['S1 Year-5 SOM','$581k','Forecast revenue, not market size'],['S2 global TAM / GovTech SAM','$1.11tn / $14bn','Founder estimates; citations unconfirmed'],['S2 regional SOM','$325m','Includes $300m transaction upside; unvalidated']],sub='No external market verification commissioned; estimates are not presented as independently substantiated.',size=18)
blocks('Compete on traceable orchestration and local adoption','Competitive positioning',[('Alternative workflows','Chat groups, spreadsheets and municipal CRMs already exist. Integrate with official systems rather than claiming to replace them.'),('Proposed differentiation','Preserved source history, report-to-incident linking, low-bandwidth intake and evidence-aware closure across channels.'),('Defensibility test','Measure onboarding speed, data quality, official-reference coverage, renewal and switching costs. No competitor benchmark supplied.')])
blocks('Win a bounded scope before scaling the footprint','Go-to-market',[('Acquire','Founder access to local representatives; map sponsor CSI teams, authorised municipal buyers and planning institutions.'),('Convert','Package one decision-use case, a paid statement of work, data authority and a named budget owner. Treat activation separately from collection.'),('Retain','Deliver a comparable evidence pack, record buyer use, measure support effort and test renewal before adding the next cohort.')])
table('First-year delivery follows four evidence gates','Execution',['Timing','Delivery','Exit gate'],[['Days 0–30','Define ward schema, scope and buyer','Named owner, authority and budget'],['Days 31–90','Ingest / profile first cohort','Traceable imports and quality review'],['Months 4–6','Deliver evidence; test paid use','Accepted output and actual collection'],['Months 7–12','Repeat onboarding; test renewal','No bespoke rebuild; viable delivery margin']],sub='Ward 115 plus up to four additional wards; free users are not counted as paying seats.',size=19)
blocks('Three horizons; regional launches remain conditional','Growth horizons',[('Y1 / Establish','Stabilise the reference product and validate the five-ward cohort. Establish customer, delivery and collection economics.'),('Y2–3 / Monetise','Develop recurring sponsor/institutional support and governed data products. CivicRewards is an optional upcoming pilot.'),('Y4–5 / Regional','Track A: Namibia, Botswana, Kenya. Track B: Nigeria, Ethiopia. Require funded local partners, lawful data access and viable collections.')])
table('Hiring must follow revenue and funded delivery capacity','Team',['Capability','Y1 / Y3 / Y5','Accountability'],[['Founders','2 / 2 / 2','Thami: institutional BD; Gugu: technical execution'],['Data / product','2 / 2 / 3','Reliable ingestion, models, security'],['Customer success','1 / 1 / 2','Training, support and retention'],['Sales / partnerships','0 / 1 / 2','Buyer pipeline and renewals'],['Governance / audit','0 / 1 / 1','Rights, data protection and assurance']],sub='S1 reference FTEs: 5 / 6 / 7 / 9 / 10; blended $45k/year. Founder experience: S2, 10+ and 28+ years.',size=17)
blocks('Governed data is a prerequisite, not a monetisation shortcut','Governance',[('Lawful use','Confirm controller/processor roles, collection basis, group-history authority and purpose limits. Obtain qualified POPIA advice.'),('Protected outputs','Minimise identifiers, suppress small groups and assess re-identification risk. Anonymisation claims require testing and contractual controls.'),('Country launch','Review local data law, hosting and transfers, procurement eligibility and political neutrality. Raw community data is not assumed saleable.')])
table('Principal risks have explicit owners and responses','Risk register',['Risk','Response / decision gate','Owner'],[['Procurement & late payment','Funded scope; stage invoices; monitor collections','Commercial founder'],['Unverified outcomes / poor data','Evidence review, independent feedback, reopen cases','Product / success'],['Privacy & data licensing rights','Legal review and aggregate release controls','Governance'],['Funding / hiring ahead of sales','Tranche approvals and monthly cash forecast','Founders'],['Rewards / payment complexity','Separate optional pilot; exclude unvalidated margins','Product / partners']],size=18)
blocks('Financial projections are planning cases, not actual results','Financial assumptions',[('Source reference','S1 prices: $10/month councillor, $50/month business, $10k/year municipality. Annual subscriptions use opening/closing customer averages.'),('Cost basis','$45k/FTE salary, $1.8k/FTE premises, $1.2k/FTE tooling. $0.02 parsing unit implies 10k → 500k annual processing units.'),('Additional assumptions','Illustrative 30-day revenue receivables, no payables, three-year straight-line hardware depreciation, no debt, interest or dividends. USD; no FX assumption.')])
table('Customer counts drive the paid-seat reference case','Financial drivers',['Year-end units','Y1','Y2','Y3','Y4','Y5'],[['Councillors','50','150','400','900','1,600'],['Businesses','5','15','40','90','180'],['Municipal contracts','0','1','3','7','15'],['Total FTEs','5','6','7','9','10'],['Opening customer baseline','0','Prior Y1','Prior Y2','Prior Y3','Prior Y4']],sub='S1 model reproduced as an alternative; paid councillor assumptions conflict with S3 free-core strategy.',size=18)
table('Five-year revenue reaches $581k in the source case','Revenue forecast',H,[row('Councillor subscriptions',coun),row('Business subscriptions',biz),row('Municipal contracts',mun),row('Data licensing',data),row('Civic advertising',ads),row('Total revenue',rev)],sub='US$000; source assumptions, not contracted sales. Data and advertising lack supplied price × volume drivers.',size=18)
s=new('Revenue scale and cost discipline must converge','Financial trajectory','Source-reference annual results in US$000; native editable chart.');chart(s,[('Revenue',[v/1000 for v in rev]),('EBITDA',[v/1000 for v in ebitda])]);text(s,'$581k',9.2,3.2,3.3,.9,42,True,'blue');text(s,'Year-5 annual revenue',9.2,4.1,3.3,.55,19);text(s,'$5k',9.2,5,3.3,.8,36,True,'green');text(s,'Year-5 EBITDA — thin buffer',9.2,5.8,3.3,.6,18)
table('Gross margins rise to 93%; delivery costs need validation','Direct cost forecast',H,[row('LLM parsing / token API',[200,700,2000,5000,10000]),row('Cloud compute',[2400,4800,9600,18000,30000]),row('Total COGS',cogs),row('Gross profit',[r-c for r,c in zip(rev,cogs)]),['Gross margin']+[f'{(r-c)/r:.1%}' for r,c in zip(rev,cogs)]],sub='Support, security, data licensing fulfilment and payment processing may add costs not identified in S1.',size=18)
table('Personnel dominates the operating cost base','Operating expenditure',H,[row('Fully loaded personnel',salary),row('Premises',rent),row('Hardware (source expensed)',hw),row('Software tools',tools),row('GTM / travel / administration',gtm),row('Total source OPEX',opex)],sub='Hardware is reclassified as capital expenditure only in the illustrative integrated statements below.',size=18)
# Integrated statements: hardware removed from OPEX, straight-line depreciation, 30-day receivables.
dep=[sum(hw[j]/3 for j in range(max(0,i-2),i+1)) for i in range(5)]
norm=[e+h for e,h in zip(ebitda,hw)];ebit=[e-d for e,d in zip(norm,dep)]
ar=[r*30/365 for r in rev];delta=[v-(ar[i-1] if i else 0) for i,v in enumerate(ar)];cf=[e-d for e,d in zip(ebitda,delta)]
fund=[-min(v,0) for v in cf];cash=[];acc=0
for f,fin in zip(cf,fund):acc+=f+fin;cash.append(acc)
netppe=[sum(hw[:i+1])-sum(dep[:i+1]) for i in range(5)];ret=[sum(ebit[:i+1]) for i in range(5)];equity=[sum(fund[:i+1]) for i in range(5)]
table('Integrated profit & loss normalises hardware treatment','Integrated financials',H,[row('Revenue',rev),row('Gross profit',[r-c for r,c in zip(rev,cogs)]),row('Normalised EBITDA',norm),row('Depreciation',[-v for v in dep]),row('EBIT / net income',ebit),row('Source EBITDA (comparison)',ebitda)],sub='All EBIT years remain negative. Illustrative no cash tax / no deferred tax asset; loss carryforwards and tax law need confirmation.',size=18)
table('Cash flow exposes the funding requirement','Integrated financials',H,[row('Normalised EBITDA',norm),row('Increase in receivables',[-v for v in delta]),row('Hardware capex',[-v for v in hw]),row('Pre-financing cash flow',cf),row('Illustrative equity required',fund),row('Closing cash',cash)],sub='Equity plugged annually to prevent negative year-end cash, not committed capital. Excludes minimum cash reserve and intra-year timing.',size=18)
table('Illustrative balance sheet reconciles to the cash model','Integrated financials',H,[row('Cash',cash),row('Receivables',ar),row('Net hardware assets',netppe),row('Total assets',[a+b+c for a,b,c in zip(cash,ar,netppe)]),row('Cumulative equity injected',equity),row('Accumulated retained losses',ret)],sub='No debt or payables assumed. Assets = contributed equity + retained earnings; opening balance assumed zero, not historical accounts.',size=18)
blocks('A $100k raise does not fund the supplied staffing plan','Capital adequacy',[('4.8 months','$248.4k Year-1 source loss ÷ 12 = $20.7k monthly burn. $100k covers 4.83 months under an even-burn simplification.'),('$981k','Peak cumulative source EBITDA deficit through Y4, before working capital or cash buffers. $881k above the proposed first tranche.'),(f'${sum(fund)/1000:,.1f}k','Illustrative total equity to maintain non-negative year-end cash with 30-day receivables. Additional reserves and monthly peaks remain unfunded.')])
table('Twelve-month bridge makes the early cash gap visible','Monthly planning',['Month band','Revenue / month','Net burn / month','Cash after band'],[['M1–M3','$642','$20,700','$37,900'],['M4–M6','$642','$20,700','($24,200)'],['M7–M9','$642','$20,700','($86,300)'],['M10–M12','$642','$20,700','($148,400)']],sub='S1 Year-1 totals divided evenly by 12; $100k opening cash. Excludes receivables timing; illustrative bridge, not a seasonal acquisition forecast.',size=20)
table('Free-core strategy requires replacement payer revenue','Commercial reconciliation',H,[row('Source-reference revenue',rev),row('Remove councillor revenue',[-v for v in coun]),row('Free-core revenue (no replacement)',free),row('Free-core EBITDA',free_ebitda),row('Annual EBITDA break-even gap',[-min(e,0) for e in free_ebitda])],sub='Retains all other unvalidated source assumptions. No sponsor prices invented. Year-5 gap $145k = 15 additional $10k contracts before delivery costs.',size=18)
blocks('Exit subscription run-rate is not total annual revenue','Unit economics',[('$450k','Year-5 exit subscription run-rate: $192k councillors + $108k businesses + $150k municipal contracts. Paid-seat alternative only.'),('$690k','Including $150k data and $90k advertising gives a total annualised illustrative run-rate, not contracted recurring ARR.'),('Still unknown','Acquisition costs, churn, retention, delivery hours, sponsor pricing and contracted data volumes. Do not publish LTV/CAC without evidence.')])
scenarios=[('Reference',rev,opex),('Revenue −25%',[r*.75 for r in rev],opex),('Revenue +25%',[r*1.25 for r in rev],opex),('Personnel +20%',rev,[o+s*.2 for o,s in zip(opex,salary)])]
rows=[]
for name,rs,ops in scenarios:
 es=[r-c-o for r,c,o in zip(rs,cogs,ops)];cum=0;peak=0
 for e in es:cum+=e;peak=max(peak,-cum)
 rows.append([name,money(rs[4]),money(es[4]),money(peak)])
table('Small revenue misses remove Year-5 break-even','Sensitivity',['Scenario','Y5 revenue $k','Y5 EBITDA $k','Peak loss funding $k'],rows,sub='Illustrative proportional revenue changes; COGS fixed conservatively. No probability weights, receivables or additional investment included.',size=20)
blocks('The source raise is a gated first tranche','Use of funds',[('40% / $40k','Product engineering and automation. Prioritise repeatable ward ingestion, linking and evidence controls.'),('35% / $35k','Secure compute and core data infrastructure. Reconcile this envelope with just $2.6k Year-1 COGS in the source model.'),('25% / $25k','Buyer development and onboarding. Rewrite the source 50 paid-councillor milestone for the free-core model.')])
blocks('Approve scale only when evidence and cash align','Investment recommendation',[('Before first close','Approve payer model, rights and source assumptions. Replace the 12–14-month runway claim with a costed, funded plan.'),('Before follow-on','Show buyer acceptance, payment, renewal interest and measured ward-level delivery costs. Negotiate scope-based financing.'),('Before regional entry','Require a local funded partner, approved data use, hosting, procurement and collection route. Avoid unfunded geography commitments.')])
table('Management decisions required before external circulation','Diligence checklist',['Decision','Required evidence','Status'],[['Commercial model','Confirm free core vs paid-seat alternative','Unresolved across sources'],['Financial baseline','Actual cash, liabilities, payroll, tax and FX','Not supplied'],['Contract pipeline','Signed scope, buyer budget, price and payment dates','Not supplied'],['Data licensing','Rights, anonymisation assessment and buyer terms','Not supplied'],['Pilot effectiveness','Independent closure and admin-time verification','Founder-reported / incomplete']],sub='Consulting-inspired presentation; not affiliated with, prepared by, or endorsed by KPMG or McKinsey.',size=18)
table('Every claim has a source or an assumption label','Source register',['ID','Document','Use'],[['S1','SIGNAL_DESK_pasted.md','Source financials; entire verbatim text in appendix'],['S2','SignalDesk-YC-PitchDeck_15Oct2026.pdf','Pilot, architecture, founders, roadmap and ask'],['S3','Signal_Desk_Full_Growth_Expansion_Plan.pptx.pdf','Free-core model, evidence, governance and regional gates'],['A1','Analyst planning assumptions','Hardware reclassification, 30-day receivables, zero opening balances'],['Limit','No independent verification','No actual accounts, committed capital or confirmed contracts']],size=17)
s=new('SIGNAL DESK','Next steps','Confirm the payer model. Fund the next evidence gate.',True);text(s,'thami@signaldesk.co.za\ngugu@signaldesk.co.za\nsignaldesk.co.za',.55,3.6,8,2,26,False,'white')
# Preserve every original line in editable source-text appendix; tables/ASCII rendered literally.
raw=Path('/mnt/user-uploads/SIGNAL_DESK_pasted.md').read_text()
# Wrap only for layout, preserve actual original text in paragraphs and speaker notes.
chunks=[];chunk=[];weight=0
for line in raw.splitlines():
 cost=max(1,math.ceil(len(line)/108))
 if weight+cost>19 and chunk:chunks.append(chunk);chunk=[];weight=0
 chunk.append(line);weight+=cost
if chunk:chunks.append(chunk)
for i,lines in enumerate(chunks):
 s=new(f'Supplied prospectus — verbatim {i+1}/{len(chunks)}','Appendix / S1','Original management text; conflicting assumptions and closing question intentionally preserved.')
 text(s,'\n'.join(lines),.6,2.95,12.05,3.9,13)
 s.notes_slide.notes_text_frame.text='Verbatim source excerpt:\n'+'\n'.join(lines)
# Full original retained also as source notes on first appendix slide.
P.core_properties.title='Signal Desk — Strategic Business Plan & Financial Projections'
P.core_properties.subject='Source-reconciled five-year business plan, October 2026'
P.core_properties.author='Signal Desk';P.core_properties.keywords='Business plan; source assumptions; editable Arial'
P.save(str(OUT/'Signal_Desk_Business_Plan.pptx'))
(ROOT/'src/assets/plan-manifest.json').write_text(json.dumps({'slides':slides,'slideCount':len(slides),'financials':{'revenue':rev,'ebitda':ebitda,'freeCoreRevenue':free,'freeCoreEbitda':free_ebitda,'peakFunding':981000,'cashFunding':round(sum(fund))}},indent=2))
print(f'Generated {len(P.slides)} slides; equity requirement ${sum(fund):,.0f}')
# Assert source arithmetic and integrated statements before QA.
assert rev==[sum(v) for v in zip(coun,biz,mun,data,ads)]
assert ebitda==[-248400,-268500,-254600,-209500,5000]
for i in range(5):assert abs(cash[i]+ar[i]+netppe[i]-equity[i]-ret[i])<.001
