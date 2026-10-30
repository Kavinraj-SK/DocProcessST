import json,os,re,struct,sys,tempfile,unicodedata,zipfile,zlib,mimetypes,time
from pathlib import Path
from dotenv import load_dotenv
try:import requests
except ImportError:requests=None
try:from docx2python import docx2python
except ImportError:sys.exit('Missing dependency. Install it with:\n    pip install docx2python --break-system-packages')
try:import docx as _pydocx
except ImportError:_pydocx=None
try:from PIL import Image as _PILImage
except ImportError:_PILImage=None
try:import imagecodecs as _imagecodecs
except ImportError:_imagecodecs=None
try:from lxml import etree as _etree
except ImportError:_etree=None
INPUT_PATH='C:\\Users\\Administrator\\Downloads\\23PT17\\zDocProcessor-main\\R2023 All Sem Missing Practicum'
OUTPUT_PATH='C:\\Users\\Administrator\\Downloads\\23PT17\\zDocProcessor-main\\R2023 All Sem Missing Practicum'
load_dotenv()
POLY_API_BASE=os.getenv('POLY_API_BASE','https://polyexam.tnedu.in/Dote/api/QuestionBank').rstrip('/')
POLY_UPLOAD_URL=os.getenv('POLY_UPLOAD_URL',POLY_API_BASE+'/UploadQuestionImages')
POLY_REFRESH_URL=os.getenv('POLY_REFRESH_URL',POLY_API_BASE+'/RefreshImageUrls')
POLY_TOKEN=os.getenv('ques_TOKEN')
UPLOAD_DELAY=float(os.getenv('UPLOAD_DELAY','1'))
UPLOAD_RETRIES=int(os.getenv('UPLOAD_RETRIES','3'))
MIN_UPLOAD_IMAGE_BYTES=int(os.getenv('MIN_UPLOAD_IMAGE_BYTES','2048'))
UPLOAD_SUPPORTED_EXTENSIONS={'.jpg','.jpeg','.png','.gif','.bmp','.webp'}
CATEGORY_HEADER_RE=re.compile('^[A-E]\\)?\\.?\\s*[A-Za-z][A-Za-z ]*SKILLS?\\s*\\(\\s*\\d+\\s*[Mm]arks?\\)?\\s*$')
STRAY_LETTER_RE=re.compile('^[A-E]\\)?\\.?\\s*$')
ROMAN_HEADING_RE=re.compile('^(VIII|III|VII|II|IV|VI|IX|I|V|X)\\b[).]?\\s*(.*)$')
SUBQ_RE=re.compile('^(\\d+)[).]\\s+(.*)$')
SUBJ_CODE_RE=re.compile('(?<!\\d)(?:[A-Za-z]{1,3}\\d{5,7}|\\d{9,10})(?!\\d)')
def extract_subj_code(stem:str)->str:
	m=SUBJ_CODE_RE.search(stem)
	if m:return m.group(0)
	return stem.split(' ',1)[0]
IMAGE_MARKER_RE=re.compile('----media/(image\\d+\\.\\w+)----')
IMAGE_ALT_MARKER_RE=re.compile('----Image alt text---->.*?<(?=----media/)',re.DOTALL)
IMAGE_ALT_MARKER_LOOSE_RE=re.compile('----Image alt text---->[^<\\n]*<?')
def strip_image_alt_markers(text:str)->str:text=IMAGE_ALT_MARKER_RE.sub('',text);return IMAGE_ALT_MARKER_LOOSE_RE.sub('',text)
LETTER_LABEL_RE=re.compile('^[A-Ea-e]\\)?\\.?$')
MARKS_TABLE_HEADER_RE=re.compile('^S\\.?\\s*NO\\.?$',re.IGNORECASE)
ROMAN_VALUES={'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9,'X':10}
ROMAN_INT_TO_STR={v:k for(k,v)in ROMAN_VALUES.items()}
def _int_to_roman(n):return ROMAN_INT_TO_STR.get(n,str(n))
def clean_cell(paragraphs):
	if paragraphs is None:return''
	if isinstance(paragraphs,str):text=paragraphs
	else:text='\n'.join(str(p)for p in paragraphs if p)
	text=strip_image_alt_markers(text.replace('\t',' ')).strip();text=_MATH_UNICODE_FIX.sub(lambda m:unicodedata.normalize('NFKC',m.group(0)),text).replace('ƒ','f')
	if text.endswith('\\'):text=text[:-1].rstrip()
	return text
def _lineage_says_table(par):lineage=getattr(par,'lineage',None);return bool(lineage)and len(lineage)>1 and lineage[1]=='tbl'
def _is_real_table(pars_table):
	for row in pars_table:
		for cell in row:
			for par in cell:
				if _lineage_says_table(par):return True
	return False

# ---------------------------------------------------------------------------
# LaTeX -> <span data-type="math-inline" ...> conversion
# Output format is matched to the frontend's cleanQuestionHtml() regex:
#   <span ... data-type="math-inline" ... data-latex="..." ...> ... </span> </span>
# The frontend decodes &amp; (globally, up front), then &quot; and &#39; inside
# the callback. It does NOT decode &lt; / &gt;, so < and > are emitted as
# \lt / \gt LaTeX commands instead.
# ---------------------------------------------------------------------------
def _clean_latex_body(latex:str)->str:
	# Normalise the LaTeX body BEFORE it goes into the span.
	latex=_fix_math_unicode(f'$${latex}$$')[2:-2]
	return re.sub(r'\s+',' ',latex).strip()
def _escape_latex_attr(latex:str)->str:
	# <, > as LaTeX commands (JS doesn't decode &lt;/&gt;).
	latex=latex.replace('<',r'\lt ').replace('>',r'\gt ')
	# & first; JS turns &amp; -> & globally, then decodes &quot; / &#39; in the callback.
	return latex.replace('&','&amp;').replace('"','&quot;').replace("'",'&#39;')
def convert_latex_to_span(text:str)->str:
	if not text:return text
	def _replace(match):
		latex=_clean_latex_body(match.group(2).strip())
		if not latex:return match.group(0)
		# The second </span> is intentional: the frontend regex requires
		# `...> </span></span>` after the data-latex attribute.
		return f'<span data-type="math-inline" data-latex="{_escape_latex_attr(latex)}"> </span></span>'
	text=re.sub('(\\$\\$)([\\s\\S]+?)(\\$\\$)',_replace,text)
	text=re.sub('(\\$)([^$\\n]+?)(\\$)',_replace,text)
	return text
def _flatten_entry(text_table):
	lines=[]
	for row in text_table:
		for cell in row:
			t=clean_cell(cell)
			if t:lines.append(t)
	return convert_latex_to_span('\n'.join(lines))
_MATH_SPLIT_RE=re.compile(r'(\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)')
# Inline styles only add borders for plain display; the rich-text editor ignores them. Set both to '' to drop them.
_TABLE_STYLE='border-collapse:collapse;width:100%'
_CELL_STYLE='border:1px solid #d1d5db;padding:4px 8px;text-align:center;vertical-align:top'
def _cell_to_html(text:str)->str:
	# One <p> per line; math segments become math-inline spans; other text is HTML-escaped.
	out=[]
	for(i,part)in enumerate(_MATH_SPLIT_RE.split(text or'')):
		if i%2==1:out.append(convert_latex_to_span(part))
		else:out.append(part.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('\n','</p><p>'))
	return'<p>'+''.join(out)+'</p>'
def _flatten_table(rows):
	# Editor-style structure: <table><tbody><tr><td><p>..</p></td>.. (single line, no <th>/<br>).
	grid=[[clean_cell(cell)for cell in row]for row in rows]
	grid=[r for r in grid if any(r)]
	if not grid:return''
	width=max(len(r)for r in grid);grid=[r+['']*(width-len(r))for r in grid]
	ts=f' style="{_TABLE_STYLE}"'if _TABLE_STYLE else'';cs=f' style="{_CELL_STYLE}"'if _CELL_STYLE else''
	html=[f'<table{ts}><tbody>']
	for r in grid:html.append('<tr>'+''.join(f'<td{cs}>{_cell_to_html(c)}</td>'for c in r)+'</tr>')
	html.append('</tbody></table>');return''.join(html)
_W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_M_NS='http://schemas.openxmlformats.org/officeDocument/2006/math'
def _m(tag):return f"{{{_M_NS}}}{tag}"
def _w(tag):return f"{{{_W_NS}}}{tag}"
_OMML_CHAR_MAP={'×':'\\times','÷':'\\div','±':'\\pm','∓':'\\mp','≤':'\\leq','≥':'\\geq','≠':'\\neq','≈':'\\approx','≡':'\\equiv','∞':'\\infty','∂':'\\partial','∇':'\\nabla','∑':'\\sum','∏':'\\prod','∫':'\\int','∬':'\\iint','∮':'\\oint','√':'\\sqrt','∠':'\\angle','⊥':'\\perp','∥':'\\parallel','→':'\\rightarrow','←':'\\leftarrow','↔':'\\leftrightarrow','⇒':'\\Rightarrow','⇐':'\\Leftarrow','⇔':'\\Leftrightarrow','∈':'\\in','∉':'\\notin','⊂':'\\subset','⊃':'\\supset','∪':'\\cup','∩':'\\cap','∅':'\\emptyset','∀':'\\forall','∃':'\\exists','¬':'\\neg','∧':'\\wedge','∨':'\\vee','°':'^{\\circ}','′':"'",'″':"''",'·':'\\cdot','…':'\\ldots','⋯':'\\cdots','⋮':'\\vdots','⋱':'\\ddots','α':'\\alpha','β':'\\beta','γ':'\\gamma','δ':'\\delta','ε':'\\epsilon','ζ':'\\zeta','η':'\\eta','θ':'\\theta','ι':'\\iota','κ':'\\kappa','λ':'\\lambda','μ':'\\mu','ν':'\\nu','ξ':'\\xi','π':'\\pi','ρ':'\\rho','σ':'\\sigma','τ':'\\tau','υ':'\\upsilon','φ':'\\phi','χ':'\\chi','ψ':'\\psi','ω':'\\omega','Α':'A','Β':'B','Γ':'\\Gamma','Δ':'\\Delta','Ε':'E','Ζ':'Z','Η':'H','Θ':'\\Theta','Ι':'I','Κ':'K','Λ':'\\Lambda','Μ':'M','Ν':'N','Ξ':'\\Xi','Ο':'O','Π':'\\Pi','Ρ':'P','Σ':'\\Sigma','Τ':'T','Υ':'\\Upsilon','Φ':'\\Phi','Χ':'X','Ψ':'\\Psi','Ω':'\\Omega'}
_FUNC_NAMES={'sin','cos','tan','cot','sec','csc','arcsin','arccos','arctan','sinh','cosh','tanh','coth','log','ln','lg','exp','det','dim','gcd','lim','max','min','sup','inf','arg'}
_ACCENT_MAP={'\u0302':'\\hat','\u0303':'\\tilde','\u0304':'\\bar','\u0307':'\\dot','\u0308':'\\ddot','\u0306':'\\breve','\u030c':'\\check','\u0301':'\\acute','\u0300':'\\grave','→':'\\vec'}
def _omml_run_text(run_el):
	parts=[]
	for t in run_el.iter(_m('t')):
		if t.text:parts.append(t.text)
	return''.join(parts)
def _latex_escape_text(s:str)->str:
	if not s:return''
	out=[]
	for ch in s:
		if ch in _OMML_CHAR_MAP:v=_OMML_CHAR_MAP[ch];out.append(v+' 'if v[-1:].isalpha()else v)
		elif ch.isspace():out.append(' ')
		elif ch in'{}%$#&_':out.append('\\'+ch)
		else:
			nk=unicodedata.normalize('NFKC',ch)
			if len(nk)==1 and ord(nk)<8192:out.append(nk)
			elif ch.isalpha()or ch.isdigit():out.append(ch)
			else:out.append(ch)
	return''.join(out)
def _wrap_if_needed(expr:str)->str:
	expr=expr.strip()
	if not expr:return''
	if len(expr)==1 or expr.startswith('{')and expr.endswith('}')or expr.startswith('\\'):return expr
	return'{'+expr+'}'
def _convert_omml_node(el)->str:
	tag=_etree.QName(el).localname
	if tag in('oMath','oMathPara'):return''.join(_convert_omml_node(c)for c in el)
	if tag=='r':return _latex_escape_text(_omml_run_text(el))
	if tag=='f':num=el.find(_m('num'));den=el.find(_m('den'));n=''.join(_convert_omml_node(c)for c in num)if num is not None else'';d=''.join(_convert_omml_node(c)for c in den)if den is not None else'';return'\\frac{'+n+'}{'+d+'}'
	if tag=='sSup':e=el.find(_m('e'));sup=el.find(_m('sup'));base=''.join(_convert_omml_node(c)for c in e)if e is not None else'';s=''.join(_convert_omml_node(c)for c in sup)if sup is not None else'';return _wrap_if_needed(base)+'^{'+s+'}'
	if tag=='sSub':e=el.find(_m('e'));sub=el.find(_m('sub'));base=''.join(_convert_omml_node(c)for c in e)if e is not None else'';s=''.join(_convert_omml_node(c)for c in sub)if sub is not None else'';return _wrap_if_needed(base)+'_{'+s+'}'
	if tag=='sSubSup':e=el.find(_m('e'));sub=el.find(_m('sub'));sup=el.find(_m('sup'));base=''.join(_convert_omml_node(c)for c in e)if e is not None else'';s=''.join(_convert_omml_node(c)for c in sub)if sub is not None else'';p=''.join(_convert_omml_node(c)for c in sup)if sup is not None else'';return _wrap_if_needed(base)+'_{'+s+'}^{'+p+'}'
	if tag=='rad':
		deg=el.find(_m('deg'));e=el.find(_m('e'));body=''.join(_convert_omml_node(c)for c in e)if e is not None else''
		if deg is not None and''.join(_convert_omml_node(c)for c in deg).strip():d=''.join(_convert_omml_node(c)for c in deg);return'\\sqrt['+d+']{'+body+'}'
		return'\\sqrt{'+body+'}'
	if tag=='nary':
		chr_el=el.find(_m('naryPr')+'/'+_m('chr'));op=None
		if chr_el is not None:op=chr_el.get(_m('val'))
		if not op:op='∫'
		latex_op=_OMML_CHAR_MAP.get(op,'\\int');sub=el.find(_m('sub'));sup=el.find(_m('sup'));e=el.find(_m('e'));out=latex_op
		if sub is not None and ''.join(_convert_omml_node(c)for c in sub).strip():out+='_{'+''.join(_convert_omml_node(c)for c in sub)+'}'
		if sup is not None and ''.join(_convert_omml_node(c)for c in sup).strip():out+='^{'+''.join(_convert_omml_node(c)for c in sup)+'}'
		if e is not None:
			body=''.join(_convert_omml_node(c)for c in e)
			if latex_op in('\\int','\\iint','\\oint'):body=re.sub(r'\s*\bd\s*([A-Za-z])\s*$',r' \\, d\1',body)
			out+=' '+body
		return out
	if tag=='d':
		beg=el.find(_m('dPr')+'/'+_m('begChr'));end=el.find(_m('dPr')+'/'+_m('endChr'));beg_ch=beg.get(_m('val'))if beg is not None else'(';end_ch=end.get(_m('val'))if end is not None else')';inner=''.join(_convert_omml_node(c)for c in el if _etree.QName(c).localname=='e');lmap={'(':'\\left(','[':'\\left[','{':'\\left\\{','⟨':'\\left\\langle','|':'\\left|'};rmap={')':'\\right)',']':'\\right]','}':'\\right\\}','⟩':'\\right\\rangle','|':'\\right|'};lb=lmap.get(beg_ch,beg_ch);rb=rmap.get(end_ch,end_ch);tall=any(k in inner for k in('\\frac','\\sum','\\prod','\\int','\\sqrt','\\begin','\\lim'))
		if not tall:lb=lb.replace('\\left','');rb=rb.replace('\\right','')
		if lb[-1:].isalpha():lb+=' '
		if rb[-1:].isalpha():rb+=' '
		return lb+inner+rb
	if tag=='func':
		fname_el=el.find(_m('fName'));e_el=el.find(_m('e'));fname=''.join(_convert_omml_node(c)for c in fname_el)if fname_el is not None else'';arg=''.join(_convert_omml_node(c)for c in e_el)if e_el is not None else'';name_clean=fname.strip().strip('\\')
		fn=fname.strip()
		if name_clean in _FUNC_NAMES:return'\\'+name_clean+' '+arg
		if fn.startswith('\\'):return fn+' '+arg
		return'\\operatorname{'+name_clean+'} '+arg
	if tag=='acc':acc_chr_el=el.find(_m('accPr')+'/'+_m('chr'));acc_chr=acc_chr_el.get(_m('val'))if acc_chr_el is not None else'\u0302';e=el.find(_m('e'));inner=''.join(_convert_omml_node(c)for c in e)if e is not None else'';cmd=_ACCENT_MAP.get(acc_chr,'\\hat');return cmd+'{'+inner+'}'
	if tag=='bar':pos=el.find(_m('barPr')+'/'+_m('pos'));pos_val=pos.get(_m('val'))if pos is not None else'top';e=el.find(_m('e'));inner=''.join(_convert_omml_node(c)for c in e)if e is not None else'';return('\\overline{'if pos_val=='top'else'\\underline{')+inner+'}'
	if tag in('limLow','limUpp'):
		e=el.find(_m('e'));lim=el.find(_m('lim'));base=''.join(_convert_omml_node(c)for c in e)if e is not None else'';low=''.join(_convert_omml_node(c)for c in lim)if lim is not None else'';b=base.strip();b=('\\'+b)if b in _FUNC_NAMES else b
		if tag=='limUpp':return'\\overset{'+low+'}{'+b+'}'
		if b.startswith('\\'):return b+'\\limits_{'+low+'}'
		return'\\underset{'+low+'}{'+b+'}'
	if tag=='m':
		rows=[]
		for mr in el.findall(_m('mr')):cells=[''.join(_convert_omml_node(c)for c in e)for e in mr.findall(_m('e'))];rows.append(' & '.join(cells))
		return'\\begin{matrix}'+' \\\\ '.join(rows)+'\\end{matrix}'
	if tag in('rPr','fPr','sSupPr','sSubPr','sSubSupPr','radPr','naryPr','dPr','funcPr','accPr','barPr','mPr','ctrlPr'):return''
	return''.join(_convert_omml_node(c)for c in el)
def omml_to_latex(omath_el)->str:
	try:latex=_convert_omml_node(omath_el).strip();latex=re.sub('\\s+',' ',latex).strip();return latex
	except Exception:return''
def _rewrite_omml_to_latex(document_xml:bytes):
	if _etree is None:return document_xml,False,0
	try:root=_etree.fromstring(document_xml)
	except Exception:return document_xml,False,0
	converted=0
	for omath in list(root.iter(_m('oMath'))):
		latex=omml_to_latex(omath)
		if not latex:continue
		parent=omath.getparent();is_display=parent is not None and _etree.QName(parent).localname=='oMathPara';wrapped='$$'+latex+'$$'if is_display else'$'+latex+'$';r=_etree.Element(_w('r'));t=_etree.SubElement(r,_w('t'));t.text=wrapped;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');p=omath.getparent()
		if p is not None:
			if is_display and p.getparent()is not None:gp=p.getparent();gp.insert(list(gp).index(p),r);p.remove(omath)
			else:idx=list(p).index(omath);p.remove(omath);p.insert(idx,r)
			converted+=1
	for para in list(root.iter(_m('oMathPara'))):
		if len(para)==0:
			p=para.getparent()
			if p is not None:p.remove(para)
	if converted==0:return document_xml,False,0
	new_xml=_etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True);return new_xml,True,converted
_W_NS_OLD='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_SUPERSCRIPT_CHARS={'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','+':'⁺','-':'⁻','−':'⁻','=':'⁼','(':'⁽',')':'⁾','a':'ᵃ','b':'ᵇ','c':'ᶜ','d':'ᵈ','e':'ᵉ','f':'ᶠ','g':'ᵍ','h':'ʰ','i':'ⁱ','j':'ʲ','k':'ᵏ','l':'ˡ','m':'ᵐ','n':'ⁿ','o':'ᵒ','p':'ᵖ','r':'ʳ','s':'ˢ','t':'ᵗ','u':'ᵘ','v':'ᵛ','w':'ʷ','x':'ˣ','y':'ʸ','z':'ᶻ'}
_SUBSCRIPT_CHARS={'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','+':'₊','-':'₋','−':'₋','=':'₌','(':'₍',')':'₎','a':'ₐ','e':'ₑ','h':'ₕ','i':'ᵢ','j':'ⱼ','k':'ₖ','l':'ₗ','m':'ₘ','n':'ₙ','o':'ₒ','p':'ₚ','r':'ᵣ','s':'ₛ','t':'ₜ','u':'ᵤ','v':'ᵥ','x':'ₓ'}
def _to_script(text:str,table:dict,fallback_marker:str)->str:
	if not text.strip():return text
	out=[]
	for ch in text:
		if ch.isspace():out.append(ch);continue
		mapped=table.get(unicodedata.normalize('NFKC',ch))
		if mapped is None:return f"{fallback_marker}({text.strip()})"
		out.append(mapped)
	return''.join(out)
def _rewrite_script_runs(document_xml:bytes):
	from lxml import etree
	def w(tag):return f"{{{_W_NS_OLD}}}{tag}"
	root=etree.fromstring(document_xml);changed=False
	for run in root.iter(w('r')):
		va=run.find(w('rPr')+'/'+w('vertAlign'))
		if va is None:continue
		kind=va.get(w('val'))
		if kind=='superscript':table,marker=_SUPERSCRIPT_CHARS,'^'
		elif kind=='subscript':table,marker=_SUBSCRIPT_CHARS,'_'
		else:continue
		for t in run.findall(w('t')):
			if t.text:
				new_text=_to_script(t.text,table,marker)
				if new_text!=t.text:t.text=new_text;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');changed=True
	if not changed:return document_xml,False
	return etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True),True
def docx_with_unicode_scripts(docx_path:Path):
	try:
		with zipfile.ZipFile(docx_path)as zin:
			original_xml=zin.read('word/document.xml');xml_after_omml,omml_changed,num_eq=_rewrite_omml_to_latex(original_xml);xml_final,script_changed=_rewrite_script_runs(xml_after_omml);changed=omml_changed or script_changed
			if not changed:return docx_path,False
			fd,tmp_name=tempfile.mkstemp(suffix='.docx');os.close(fd)
			with zipfile.ZipFile(tmp_name,'w',zipfile.ZIP_DEFLATED)as zout:
				for item in zin.infolist():data=xml_final if item.filename=='word/document.xml'else zin.read(item.filename);zout.writestr(item,data)
		if num_eq:print(f"  [INFO] Converted {num_eq} OMML equation(s) to LaTeX in {docx_path.name}")
		return Path(tmp_name),True
	except Exception as e:print(f"  [WARN] Could not normalise OMML/scripts in {docx_path.name}: {e}");return docx_path,False
def get_body_entries(docx_path:Path):
	read_path,is_temp=docx_with_unicode_scripts(docx_path)
	try:
		with docx2python(str(read_path))as doc:
			body=doc.body
			try:body_pars=doc.body_pars
			except AttributeError:body_pars=None
			entries=[]
			for(idx,text_table)in enumerate(body):
				if body_pars is not None:is_real=_is_real_table(body_pars[idx])
				else:is_real=not(len(text_table)==1 and len(text_table[0])==1)
				entries.append((is_real,text_table))
			return _split_trailing_headings(entries)
	finally:
		if is_temp:
			try:os.remove(read_path)
			except OSError:pass
def _split_trailing_headings(entries):
	expanded=[]
	for(is_real,rows)in entries:
		if is_real and len(rows)>1:
			last_row=rows[-1]
			if last_row and last_row[0]:
				cell0=clean_cell(last_row[0]).rstrip('.').strip();is_paired=len(last_row)>=2 and clean_cell(last_row[1]).strip()and len(clean_cell(last_row[1]).strip())>=15;is_bare=len(last_row)==1
				if cell0.isdigit()and cell0!='0'and(is_paired or is_bare):expanded.append((is_real,rows[:-1]));expanded.append((is_real,[last_row]));continue
		expanded.append((is_real,rows))
	return expanded
def _digit_heading_info(is_real,rows):
	if len(rows)<1 or len(rows[0])<1:return False,None,[]
	cell0=clean_cell(rows[0][0]).rstrip('.').strip()
	if not cell0.isdigit():return False,None,[]
	if cell0=='0':return False,None,[]
	if len(rows[0])>=2:
		for row in rows[1:]:
			if len(row)>=1 and clean_cell(row[0]).strip():return False,None,[]
		buffer_lines=[clean_cell(rows[0][1])]
		for extra_row in rows[1:]:
			if len(extra_row)>=2:
				t=clean_cell(extra_row[1])
				if t:buffer_lines.append(t)
		return True,cell0,buffer_lines
	if len(rows)==1 and len(rows[0])==1:return True,cell0,[]
	return False,None,[]
def _is_digit_heading(is_real,rows):is_heading,_,_=_digit_heading_info(is_real,rows);return is_heading
def _is_lettered_subanswer_table(rows):return len(rows)==1 and len(rows[0])==2 and bool(LETTER_LABEL_RE.match(clean_cell(rows[0][0]).strip()))and bool(clean_cell(rows[0][1]).strip())
def _is_stop_marker(rows):
	if not rows or not rows[0]:return False
	first_cell=clean_cell(rows[0][0]).strip()
	if first_cell.upper().startswith('ALLOCATION OF MARKS'):return True
	if MARKS_TABLE_HEADER_RE.match(first_cell):return True
	return False
def collect_digit_blocks(entries):
	blocks=[];i=0;n=len(entries)
	while i<n:
		is_real,rows=entries[i];is_heading,digit,initial_buffer=_digit_heading_info(is_real,rows)
		if is_heading:
			buffer_lines=list(initial_buffer);i+=1
			while i<n:
				is_real2,rows2=entries[i];is_heading2,_,_=_digit_heading_info(is_real2,rows2)
				if is_heading2:break
				if _is_stop_marker(rows2):i+=1;break
				if not is_real2:
					flat=_flatten_entry(rows2)
					if flat.strip()and flat.strip()!='0':buffer_lines.append(flat)
					i+=1
				elif _is_lettered_subanswer_table(rows2):label=clean_cell(rows2[0][0]).strip();desc=clean_cell(rows2[0][1]).strip();buffer_lines.append(f"{label} {desc}");i+=1
				else:
					flat_table=_flatten_table(rows2)
					if flat_table.strip():buffer_lines.append(flat_table)
					i+=1
			blocks.append((digit,'\n'.join(buffer_lines)))
		else:i+=1
	return blocks
def parse_digit_block(digit,raw_text,flags):
	lines=[l.rstrip()for l in raw_text.split('\n')];sections=[];preamble_lines=[];just_saw_category_header=False;seen_any_roman=False;next_position=1
	for line in lines:
		stripped=line.strip()
		if not stripped:continue
		if CATEGORY_HEADER_RE.match(stripped):just_saw_category_header=True;continue
		if STRAY_LETTER_RE.match(stripped):continue
		m=ROMAN_HEADING_RE.match(stripped)
		if m:
			token,rest=m.group(1),m.group(2);accept=False
			if token!='I':accept=True
			elif not seen_any_roman or just_saw_category_header:accept=True
			if accept:
				is_expected_reset=not seen_any_roman or just_saw_category_header;token_value=ROMAN_VALUES[token]
				if not is_expected_reset and token_value!=next_position:flags.append({'digit':digit,'line':stripped,'reason':f"Roman numeral '{token}' found where '{_int_to_roman(next_position)}' was expected next - possible typo in the source document's own numbering (not a parsing ambiguity). Verify the sequence is correct."})
				seen_any_roman=True;just_saw_category_header=False;next_position=token_value+1;sections.append({'roman':token,'lines':[rest]if rest else[]});continue
			else:flags.append({'digit':digit,'line':stripped,'reason':"Line starts with 'I' but is not the first Roman heading and doesn't follow a category header - treated as body text, not a new subsection. Verify this is correct."})
		else:
			am=SUBQ_RE.match(stripped)
			if am and just_saw_category_header:roman_token=_int_to_roman(next_position);seen_any_roman=True;just_saw_category_header=False;next_position+=1;sections.append({'roman':roman_token,'lines':[am.group(2)]if am.group(2)else[]});flags.append({'digit':digit,'line':stripped,'reason':f"Section heading used Arabic numeral '{am.group(1)}' instead of a Roman numeral in the source document - normalized to '{roman_token}' since it directly follows a category header. Verify this is correct."});continue
		just_saw_category_header=False
		if sections:sections[-1]['lines'].append(stripped)
		else:preamble_lines.append(stripped)
	if not sections:description='\n'.join(preamble_lines).strip();return{'Qno':digit,'Description':convert_latex_to_span(description)}
	sub_divisions=[]
	for sec in sections:
		roman=sec['roman'];body_lines=sec['lines'];sub_questions=[];description_lines=[]
		for bl in body_lines:
			sm=SUBQ_RE.match(bl)
			if sm:sub_questions.append({'Qno':sm.group(1),'ParentQuestionId':roman,'Description':convert_latex_to_span(sm.group(2).strip())})
			elif not sub_questions:description_lines.append(bl)
			elif sub_questions:sub_questions[-1]['Description']+='\n'+bl
		desc_text=convert_latex_to_span('\n'.join(description_lines).strip());node={'Qno':roman,'ParentQuestionId':digit,'Description':desc_text}
		if sub_questions:node['SubQuestions']=sub_questions
		sub_divisions.append(node)
	return{'Qno':digit,'SubDivisions':sub_divisions}
def extract_questions_from_docx(docx_path:Path):
    entries=get_body_entries(docx_path);blocks=collect_digit_blocks(entries);flags=[]
    questions=[parse_digit_block(digit,raw_text,flags)for(digit,raw_text)in blocks]if blocks else[]
    # Questions as rows of one numbered table ("1. | text", "2. | text", ...): the standard parser
    # only sees the lone last row of such a table, so prefer this reading when it finds more.
    row_questions=_parse_numbered_row_tables(entries)
    if len(row_questions)>len(questions):
        return row_questions,[{'reason':"Questions are the numbered rows of one table ('1. | text'); parsed one question per row. Please review the questions before uploading."}]
    if questions:return questions,flags
    # No "1 | question text" headings: try the fallback layouts below (only used in this case).
    questions,layout=_parse_fallback_layouts(entries)
    if questions:
        flags.append({'reason':f"No standard question headings found; parsed with the fallback layout '{layout}'. Please review the questions before uploading."})
    return questions,flags

# ---------------------------------------------------------------------------
# Fallback layouts. Used only when collect_digit_blocks() finds nothing.
#   1. practical blocks : "PRACTICAL PART / Question ..." paragraphs; everything after an
#                         "OUTPUT" line (the student's answer-sheet tables) is dropped
#   2. numbered list    : "1) ...  2) ..." lines (typically all inside one table cell)
#   3. heading items    : "(i). ..." headings with "(a) ..." sub-items and image lines
# ---------------------------------------------------------------------------
_STOP='\x00STOP'
_NUM_Q_RE=re.compile(r'^\s*(\d{1,3})\s*[).]\s+(\S.*)$')
_PRACTICAL_MARK_RE=re.compile(r'^\s*(?:[ivxIVX]+\s*[.)])?\s*PRACTICAL\s*PART\b',re.I)
_OUTPUT_RE=re.compile(r'^\s*OUTPUT\b',re.I)
_QUESTION_WORD_RE=re.compile(r'^\s*Question\s*:?\s*$',re.I)
_ROMAN_ITEM_RE=re.compile(r'^\s*\(\s*(?:i{1,3}|iv|v|vi{1,3}|ix|x)\s*\)\s*\.?\s*(\S.*)$',re.I)
_LETTER_ITEM_RE=re.compile(r'^\s*\(?\s*([a-z])\s*\)\s*(\S.*)$',re.I)
_IMG_LINE_RE=re.compile(r'^\s*----media/[A-Za-z0-9_.-]+----\s*$')

def _entry_lines(is_real,rows):
    """One body entry -> list of text lines. A real data table becomes a single HTML line."""
    filled=[c for c in(clean_cell(c)for r in rows for c in r)if c]
    if not filled:return[]
    if is_real and not(len(rows)==1 and len(filled)==1):
        html=_flatten_table(rows)
        return[html]if html.strip()else[]
    return convert_latex_to_span('\n'.join(filled)).split('\n')

def _all_lines(entries):
    lines=[]
    for is_real,rows in entries:
        if _is_stop_marker(rows):lines.append(_STOP)
        lines.extend(_entry_lines(is_real,rows))
    return lines

def _parse_practical_blocks(lines):
    qs=[];cur=None;in_output=False
    for ln in lines:
        # stop markers are ignored here: answer-sheet tables also have an "S.No." header,
        # and the "OUTPUT" line already marks where each question ends
        if ln==_STOP:continue
        if _PRACTICAL_MARK_RE.match(ln):
            cur={'Qno':str(len(qs)+1),'Description':''};qs.append(cur);in_output=False
            rest=_PRACTICAL_MARK_RE.sub('',ln).strip()
            if rest and not _QUESTION_WORD_RE.match(rest):cur['Description']=rest
            continue
        if cur is None or in_output:continue
        if _OUTPUT_RE.match(ln):in_output=True;continue
        if _QUESTION_WORD_RE.match(ln)and not cur['Description']:continue
        cur['Description']+=('\n'if cur['Description']else'')+ln.strip()
    return[q for q in qs if q['Description'].strip()]

def _parse_numbered_lines(lines):
    qs=[];cur=None
    for ln in lines:
        if ln==_STOP:
            if cur is not None:break
            continue
        m=None if ln.lstrip().startswith('<')else _NUM_Q_RE.match(ln)
        # a new question only if the number is not lower than the next expected one,
        # so a stray "2)" inside question 7 stays part of question 7
        if m and int(m.group(1))>=len(qs)+1:
            cur={'Qno':m.group(1),'Description':m.group(2).strip()};qs.append(cur)
        elif cur is not None:
            cur['Description']+='\n'+ln.strip()
    return qs if len(qs)>=2 else[]

def _parse_heading_items(lines):
    qs=[];cur=None;sub=None
    for ln in lines:
        if ln==_STOP:
            if cur is not None:break
            continue
        s=ln.strip()
        if _ROMAN_ITEM_RE.match(s):
            cur={'Qno':str(len(qs)+1),'Description':s};qs.append(cur);sub=None;continue
        if cur is None:continue
        ml=None if _IMG_LINE_RE.match(s)or s.startswith('<')else _LETTER_ITEM_RE.match(s)
        if ml:
            subs=cur.setdefault('SubQuestions',[]);expected=chr(ord('a')+len(subs))
            if ml.group(1).lower()>=expected:
                sub={'Qno':ml.group(1).lower(),'ParentQuestionId':cur['Qno'],'Description':s};subs.append(sub);continue
        target=sub if sub is not None else cur
        target['Description']+='\n'+s
    return qs if len(qs)>=2 else[]

_ROWNUM_RE=re.compile(r'^\s*(\d{1,3})\s*[.)]?\s*$')

def _question_rows(rows):
    """[(number, text), ...] if EVERY non-empty row of this table is "<number> | <question text>", else None.
    Marks tables are rejected: they have a header row, a short text cell or a numeric marks column."""
    out=[]
    for r in rows:
        cells=[clean_cell(c)for c in r]
        if not any(cells):continue
        if len(cells)<2:return None
        m=_ROWNUM_RE.match(cells[0]);text=cells[1]
        if not m or len(text)<12:return None
        if len(cells)>=3 and cells[2].replace('.','').strip().isdigit():return None
        out.append((m.group(1),text))
    return out or None

def _parse_numbered_row_tables(entries):
    qs=[]
    for is_real,rows in entries:
        if _is_stop_marker(rows):
            if qs:break
            continue
        if not is_real:continue
        qr=_question_rows(rows)
        if not qr:continue
        for num,text in qr:qs.append({'Qno':num,'Description':convert_latex_to_span(text)})
    return qs if len(qs)>=2 else[]

def _parse_fallback_layouts(entries):
    lines=_all_lines(entries)
    for name,fn in(('practical blocks',_parse_practical_blocks),('numbered list',_parse_numbered_lines),('heading items',_parse_heading_items)):
        qs=fn(lines)
        if qs:return qs,name
    return[],None
def has_omml_equations(docx_path:Path)->bool:
	if _pydocx is None:return False
	try:d=_pydocx.Document(str(docx_path));return'<m:oMath'in d.element.xml
	except Exception:return False
def has_embedded_images(docx_path:Path)->bool:
	if _pydocx is None:return False
	try:d=_pydocx.Document(str(docx_path));xml=d.element.xml;return'<w:drawing'in xml or'<a:blip'in xml
	except Exception:return False
_MATH_UNICODE_FIX=re.compile('[\\U0001D400-\\U0001D7FF]')
_TRIG_NAMES=['sin','cos','tan','cot','sec','csc','arcsin','arccos','arctan','sinh','cosh','tanh','log','ln','exp']
def _fix_latex_in_math(text:str)->str:
	if not text:return text
	def _fix_math_block(m):
		opener,body,closer=m.group(1),m.group(2),m.group(3)
		for name in _TRIG_NAMES:body=re.sub(rf"(?<!\\)\b{name}(?=\s*[A-Za-z0-9(])",rf"\\{name} ",body)
		body=re.sub(r'(?<!\\)\blim\s*([A-Za-z])\s*(?:\\rightarrow|\\to|→)\s*(-?[A-Za-z0-9.]+)',r'\\lim\\limits_{\1\\rightarrow \2}',body);body=re.sub(r'(?<!\\)\blim\s*([A-Za-z])(?=\s*\\?(?:rightarrow|to|→))',r'\\lim_{\1}',body);body=re.sub('(?<!\\\\)\\blim\\s+([A-Za-z])','\\\\lim_{\\1}',body);opens=body.count('{');closes=body.count('}')
		if opens>closes:body+='}'*(opens-closes)
		return opener+body+closer
	text=re.sub('(\\$\\$)([\\s\\S]+?)(\\$\\$)',_fix_math_block,text);text=re.sub('(\\$)([^$\\n]+?)(\\$)',_fix_math_block,text);return text
def _fix_math_unicode(text:str)->str:
	if not text:return text
	text=_MATH_UNICODE_FIX.sub(lambda m:unicodedata.normalize('NFKC',m.group(0)),text)
	# (removed: text.replace('">','"') -- it corrupted HTML attributes)
	text=re.sub('\\\\~\\{([^}]*)\\}','\\\\tilde{\\1}',text);text=re.sub("\\\\'\\{([^}]*)\\}",'\\\\acute{\\1}',text);text=re.sub('\\\\`\\{([^}]*)\\}','\\\\grave{\\1}',text);text=re.sub('\\\\\\^\\{([^}]*)\\}','\\\\hat{\\1}',text);text=re.sub('\\\\"\\{([^}]*)\\}','\\\\ddot{\\1}',text)
	return _fix_latex_in_math(text)

def _apply_math_cleanup(questions):
	# Cleanup now happens inside convert_latex_to_span; nothing left to do here.
	return
def extract_embedded_media(docx_path:Path,out_dir:Path):
	media_dir=out_dir/f"{docx_path.stem}_media";media_dir.mkdir(parents=True,exist_ok=True);extracted=[]
	try:
		with zipfile.ZipFile(docx_path)as archive:
			for name in archive.namelist():
				if not name.startswith('word/media/')or name.endswith('/'):continue
				target=media_dir/Path(name).name
				with archive.open(name)as src,open(target,'wb')as dst:dst.write(src.read())
				extracted.append(target.name)
	except Exception:return None,[]
	return media_dir.name,extracted
def _pad_jpeg(data:bytes,target_size:int)->bytes:
	needed=target_size-len(data)
	if needed<=0 or not data.endswith(b'\xff\xd9'):return data
	overhead=4;payload_len=max(needed-overhead,0);segments=b'';remaining=payload_len
	while remaining>0:chunk=min(remaining,65533);segments+=b'\xff\xfe'+struct.pack('>H',chunk+2)+b'\x00'*chunk;remaining-=chunk
	return data[:-2]+segments+data[-2:]
def _pad_png(data:bytes,target_size:int)->bytes:
	needed=target_size-len(data)
	if needed<=0 or not data.endswith(b'IEND\xaeB`\x82'):return data
	overhead=20;payload_len=max(needed-overhead,1);chunk_type=b'tEXt';chunk_data=b'Padding\x00'+b'0'*payload_len;chunk=struct.pack('>I',len(chunk_data))+chunk_type+chunk_data+struct.pack('>I',zlib.crc32(chunk_type+chunk_data)&4294967295);iend_start=len(data)-12;return data[:iend_start]+chunk+data[iend_start:]
def pad_image_to_min_size(path:Path,min_bytes:int=MIN_UPLOAD_IMAGE_BYTES)->bool:
	try:data=path.read_bytes()
	except Exception:return False
	if len(data)>=min_bytes:return True
	if data.startswith(b'\xff\xd8\xff'):padded=_pad_jpeg(data,min_bytes)
	elif data.startswith(b'\x89PNG\r\n\x1a\n'):padded=_pad_png(data,min_bytes)
	else:return False
	if len(padded)<min_bytes:return False
	path.write_bytes(padded);return True
def convert_unsupported_image_format(path:Path):
	ext=path.suffix.lower()
	if ext in UPLOAD_SUPPORTED_EXTENSIONS:return
	out_path=path.with_suffix('.png')
	def _save_array_as_png(arr):
		img=_PILImage.fromarray(arr)
		if img.mode not in('RGB','RGBA','L'):img=img.convert('RGB')
		img.save(out_path,'PNG');return out_path
	pil_error=None
	if _PILImage is not None:
		try:
			img=_PILImage.open(path);img.load()
			if img.mode not in('RGB','RGBA','L'):img=img.convert('RGB')
			img.save(out_path,'PNG');return out_path
		except Exception as e:pil_error=e
	if _imagecodecs is not None and _PILImage is not None:
		try:arr=_imagecodecs.imread(path.read_bytes());return _save_array_as_png(arr)
		except Exception as e:
			if ext in('.emf','.wmf','.svg'):print(f"    [WARN] {path.name} is a vector format ('{ext}') -- neither Pillow nor imagecodecs can rasterize vector drawings; a rendering engine like LibreOffice or Inkscape would be needed -- skipping");return
			print(f"    [WARN] couldn't convert unsupported format {path.name} ({ext}): {e}");return
	if _PILImage is None:print(f"    [WARN] {path.name} has unsupported format '{ext}' and Pillow isn't installed -- install with:\n        pip install Pillow imagecodecs --break-system-packages")
	else:print(f"    [WARN] {path.name} has unsupported format '{ext}'; Pillow couldn't read it ({pil_error}) and 'imagecodecs' isn't installed -- install with:\n        pip install imagecodecs --break-system-packages")

def load_existing_media_map(json_path:Path)->dict:
	if not json_path.exists():return{}
	try:
		with open(json_path,'r',encoding='utf-8')as f:data=json.load(f)
	except Exception:return{}
	mapping={}
	for entry in data.get('_embedded_media')or[]:
		fname=entry.get('fileName')
		if fname and entry.get('filePath'):mapping[fname]={'filePath':entry['filePath'],'imageUrl':entry.get('imageUrl','')}
	return mapping

def refresh_baked_urls(questions,url_by_basename:dict):
	if not url_by_basename:return
	patterns=[(re.compile('src="[^"]*'+re.escape(basename)+'[^"]*"'),f'src="{fresh_url}"')for(basename,fresh_url)in url_by_basename.items()]
	def process_node(node):
		if node.get('Description')is not None:
			text=node['Description']
			for(pattern,replacement)in patterns:text=pattern.sub(replacement,text)
			node['Description']=text
		for child in node.get('SubDivisions',[])or[]:process_node(child)
		for child in node.get('SubQuestions',[])or[]:process_node(child)
	for q in questions:process_node(q)

# =====================================================================================
# PASTE THIS WHOLE FILE'S CONTENT into doc_to_json.py, just above `def run_pipeline`.
# Then delete the OLD versions of: upload_images_to_s3, fetch_view_url, refresh_json_file,
# substitute_image_placeholders, process_file  (and the SFM_* lines + BEARER_TOKEN at the top).
# Keep: load_existing_media_map, refresh_baked_urls.
# Top of file must also have:  import mimetypes   (add to the first import line)
# and these config lines (replace the SFM_* ones):
#   POLY_API_BASE=os.getenv('POLY_API_BASE','https://polyexam.tnedu.in/Dote/api/QuestionBank').rstrip('/')
#   POLY_UPLOAD_URL=os.getenv('POLY_UPLOAD_URL',POLY_API_BASE+'/UploadQuestionImages')
#   POLY_REFRESH_URL=os.getenv('POLY_REFRESH_URL',POLY_API_BASE+'/RefreshImageUrls')
#   POLY_TOKEN=os.getenv('ques_TOKEN')
# =====================================================================================
# ======================= POLY IMAGE HANDLING (replaces all SFM code) =======================
# Config (put near the other env lines at the top of the file):
#   POLY_API_BASE=os.getenv('POLY_API_BASE','https://polyexam.tnedu.in/Dote/api/QuestionBank').rstrip('/')
#   POLY_UPLOAD_URL=os.getenv('POLY_UPLOAD_URL',POLY_API_BASE+'/UploadQuestionImages')
#   POLY_REFRESH_URL=os.getenv('POLY_REFRESH_URL',POLY_API_BASE+'/RefreshImageUrls')
#   POLY_TOKEN=os.getenv('ques_TOKEN')
IMG_STYLE='max-width: 350px; max-height: 300px; width: auto; height: auto; object-fit: contain; display: block; margin: 8px auto;'

def _poly_headers():
    return {'Authorization': f'Bearer {POLY_TOKEN}'} if POLY_TOKEN else {}

def fetch_view_urls(keys):
    """Batch RefreshImageUrls call. Returns {filePath: https url}; keys the server could not sign are left out."""
    keys = sorted({k for k in keys if k})
    if not keys or requests is None:
        return {}
    fresh = {}
    for i in range(0, len(keys), 25):
        chunk = keys[i:i + 25]
        try:
            resp = requests.post(POLY_REFRESH_URL, headers=_poly_headers(), json=chunk, timeout=120)
            body = resp.json()
        except Exception as e:
            print(f"    [WARN] RefreshImageUrls failed for {len(chunk)} key(s): {e}")
            continue
        for img in body.get('images') or []:
            fp, url = img.get('filePath'), img.get('imageUrl') or ''
            if fp and url.lower().startswith('http'):
                fresh[fp] = url
            elif fp:
                print(f"    [WARN] no fresh URL for {fp}: {url[:150]}")
    return fresh

def fetch_view_url(file_path_key: str) -> str:
    return fetch_view_urls([file_path_key]).get(file_path_key, '')

def upload_images_to_s3(image_paths, docx_stem):
    """Upload through the poly UploadQuestionImages endpoint (same path the editor uses)."""
    if not image_paths or requests is None:
        return {}
    mapping = {}
    for local_path in image_paths:
        upload_path = local_path
        if local_path.suffix.lower() not in UPLOAD_SUPPORTED_EXTENSIONS:
            converted = convert_unsupported_image_format(local_path)
            if converted is None:
                continue
            upload_path = converted
        if not pad_image_to_min_size(upload_path):
            print(f"    [WARN] {upload_path.name} is under {MIN_UPLOAD_IMAGE_BYTES} bytes and couldn't be padded")
        unique_name = f"{docx_stem}_{upload_path.name}"
        mime = mimetypes.guess_type(upload_path.name)[0] or 'application/octet-stream'
        # Uploads go out one after another. The server's permission check sometimes answers
        # "No write permission" / "Permission service failed" when called in quick succession,
        # so wait a little between images and retry a few times when that is the answer.
        time.sleep(UPLOAD_DELAY)
        body, resp, failed = {}, None, False
        for attempt in range(1, UPLOAD_RETRIES + 1):
            try:
                with open(upload_path, 'rb') as fh:
                    resp = requests.post(POLY_UPLOAD_URL, files=[('files', (unique_name, fh, mime))],
                                         headers=_poly_headers(), timeout=120)
                try:
                    body = resp.json()
                except ValueError:
                    body = {}
            except Exception as e:
                print(f"    [WARN] image upload failed for {upload_path.name}: {e}")
                failed = True
                break
            got_path = bool((body.get('images') or [{}])[0].get('filePath')) if body.get('images') else False
            if got_path or 'permission' not in (resp.text or '').lower() or attempt == UPLOAD_RETRIES:
                break
            print(f"    [RETRY] {upload_path.name}: permission error, attempt {attempt}/{UPLOAD_RETRIES}")
            time.sleep(UPLOAD_DELAY + 2 * attempt)
        if failed:
            continue
        images = body.get('images') or []
        if not images or not images[0].get('filePath'):
            print(f"    [WARN] upload gave no filePath for {upload_path.name}: {resp.status_code} {resp.text[:300]}")
            continue
        key = images[0]['filePath']
        url = images[0].get('imageUrl') or ''
        if not url.lower().startswith('http'):
            url = fetch_view_url(key)
        mapping[local_path.name] = {'filePath': key, 'imageUrl': url}
    return mapping

def substitute_image_placeholders(questions, url_by_filename, key_by_filename=None):
    key_by_filename = key_by_filename or {}
    def repl(m):
        fname = m.group(1)
        key = key_by_filename.get(fname)
        if not key:
            return m.group(0)
        url = url_by_filename.get(fname, '')
        return f'<img src="{url}" data-file-path="{key}" alt="{fname}" style="{IMG_STYLE}" />'
    def walk(node):
        if node.get('Description') is not None:
            node['Description'] = IMAGE_MARKER_RE.sub(repl, strip_image_alt_markers(node['Description']))
        for child in (node.get('SubDivisions') or []) + (node.get('SubQuestions') or []):
            walk(child)
    for q in questions:
        walk(q)

def refresh_json_file(json_path: Path) -> bool:
    """--refresh mode: refresh _embedded_media URLs and the signed src inside every question, via poly."""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"  [ERROR] could not read {json_path.name}: {e}")
        return False
    media = data.get('_embedded_media') or []
    keys = [m['filePath'] for m in media if m.get('filePath')]
    if not keys:
        print(f"  [SKIP] {json_path.name}: no embedded media to refresh")
        return False
    fresh = fetch_view_urls(keys)
    if not fresh:
        print(f"  [SKIP] {json_path.name}: no URLs could be refreshed")
        return False
    url_by_basename = {}
    for m in media:
        if m.get('filePath') in fresh:
            m['imageUrl'] = fresh[m['filePath']]
            url_by_basename[Path(m['filePath']).name] = m['imageUrl']
    refresh_baked_urls(data.get('questions', []), url_by_basename)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  [OK] {json_path.name}: refreshed {len(url_by_basename)}/{len(keys)} image URL(s)")
    return True

def process_file(docx_path: Path, out_dir: Path):
    has_equations = has_omml_equations(docx_path)
    try:
        questions, flags = extract_questions_from_docx(docx_path)
    except Exception as e:
        print(f"  [ERROR] Failed to process {docx_path.name}: {e}")
        return 'error'
    if not questions:
        print(f"  [SKIP] {docx_path.name}: no questions found")
        return 'no-questions'
    _apply_math_cleanup(questions)
    full_text = json.dumps(questions, ensure_ascii=False)
    media_dir_name, embedded_media = None, []
    if has_embedded_images(docx_path) or IMAGE_MARKER_RE.search(full_text):
        media_dir_name, embedded_media = extract_embedded_media(docx_path, out_dir)
    uploaded_media = []
    if embedded_media and media_dir_name:
        media_dir = out_dir / media_dir_name
        existing_out_path = out_dir / (docx_path.stem + '.json')
        already_uploaded = load_existing_media_map(existing_out_path)
        # reuse an earlier upload ONLY if poly can still serve it; otherwise upload again
        fresh_old = fetch_view_urls([already_uploaded[f]['filePath'] for f in embedded_media if f in already_uploaded])
        url_map, to_upload = {}, []
        for fname in embedded_media:
            old = already_uploaded.get(fname)
            if old and fresh_old.get(old['filePath']):
                url_map[fname] = {'filePath': old['filePath'], 'imageUrl': fresh_old[old['filePath']]}
                print(f"    [REUSE] {fname} already uploaded at {old['filePath']}")
            else:
                to_upload.append(fname)
        if to_upload:
            image_paths = [media_dir / fname for fname in to_upload]
            url_map.update(upload_images_to_s3(image_paths, docx_path.stem))
        for fname in embedded_media:
            entry = {'fileName': fname}
            if fname in url_map:
                entry.update(url_map[fname])
            uploaded_media.append(entry)
        url_by_filename = {f: i.get('imageUrl', '') for f, i in url_map.items() if i.get('filePath')}
        key_by_filename = {f: i['filePath'] for f, i in url_map.items() if i.get('filePath')}
        if key_by_filename:
            substitute_image_placeholders(questions, url_by_filename, key_by_filename)
    subj_code = extract_subj_code(docx_path.stem)
    output = {'subj_code': subj_code, 'questions': questions}
    if flags:
        output['_review_flags'] = flags
    if has_equations:
        output['_review_flags'] = output.get('_review_flags', []) + [{
            'file': docx_path.name,
            'reason': 'document contained embedded equation object(s); converted to inline LaTeX for frontend rendering'}]
    if media_dir_name:
        output['_media_dir'] = media_dir_name
    if uploaded_media:
        output['_embedded_media'] = uploaded_media
    elif embedded_media:
        output['_embedded_media'] = embedded_media
    out_path = out_dir / (docx_path.stem + '.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    flag_note = f", {len(flags)} line(s) flagged for review" if flags else ''
    print(f"  [OK] {docx_path.name} -> {out_path.name} ({len(questions)} question(s){flag_note})")
    return 'ok'

def run_pipeline(input_path:Path,output_path:Path)->dict:
	if not input_path.exists():raise FileNotFoundError(f"Path does not exist: {input_path}")
	if input_path.is_file():
		if input_path.suffix.lower()!='.docx':raise ValueError(f"Not a .docx file: {input_path}")
		docx_files=[input_path]
	else:docx_files=sorted(input_path.rglob('*.docx'));docx_files=[f for f in docx_files if not f.name.startswith('~$')]
	if not docx_files:raise FileNotFoundError(f"No .docx files found at: {input_path}")
	seen_stems={}
	for f in docx_files:seen_stems.setdefault(f.stem,[]).append(f)
	dupes={stem:paths for(stem,paths)in seen_stems.items()if len(paths)>1}
	if dupes:
		print('[WARNING] Duplicate filenames found across subfolders - later files will overwrite earlier JSON output:')
		for(stem,paths)in dupes.items():
			for p in paths:print(f"    {stem}.docx  <-  {p}")
	print(f"Found {len(docx_files)} .docx file(s).");output_path.mkdir(parents=True,exist_ok=True);ok_count=0;no_q_count=0;image_count=0;eqn_count=0;error_count=0
	for docx_file in docx_files:
		result=process_file(docx_file,output_path)
		if result=='ok':ok_count+=1
		elif result=='has-image':image_count+=1
		elif result=='has-equations':eqn_count+=1
		elif result=='no-questions':no_q_count+=1
		else:error_count+=1
	print(f"\nDone. {ok_count}/{len(docx_files)} file(s) converted successfully.")
	if image_count:print(f"  {image_count} file(s) skipped: contain embedded image(s) (not handled yet)")
	if eqn_count:print(f"  {eqn_count} file(s) skipped: contain embedded equation object(s) (not handled yet)")
	if no_q_count:print(f"  {no_q_count} file(s) skipped: no questions found (check structure)")
	if error_count:print(f"  {error_count} file(s) skipped: processing error")
	return{'total':len(docx_files),'ok':ok_count,'no_questions':no_q_count,'images':image_count,'equations':eqn_count,'errors':error_count,'duplicates':list(dupes.items())}
def main():
	args=sys.argv[1:]
	if args and args[0]=='--refresh':
		refresh_args=args[1:];target=Path(refresh_args[0]).expanduser().resolve()if refresh_args else Path(OUTPUT_PATH).expanduser().resolve()
		if not target.exists():sys.exit(f"Path does not exist: {target}")
		json_files=[target]if target.is_file()else sorted(target.glob('*.json'))
		if not json_files:sys.exit(f"No .json file(s) found at: {target}")
		print(f"Refreshing image URLs in {len(json_files)} file(s)...");ok_refresh=sum(1 for jf in json_files if refresh_json_file(jf));print(f"\nDone. {ok_refresh}/{len(json_files)} file(s) refreshed.");return
	input_path=Path(args[0]).expanduser().resolve()if len(args)>=1 else Path(INPUT_PATH).expanduser().resolve();output_path=Path(args[1]).expanduser().resolve()if len(args)>=2 else Path(OUTPUT_PATH).expanduser().resolve()
	try:run_pipeline(input_path,output_path)
	except(FileNotFoundError,ValueError)as e:sys.exit(str(e))
if __name__=='__main__':main()