"""Translate a small subset of Excel formulas (the Stock sheet's suggestion and max-boxes
formulas) into JavaScript expressions the Wardiya app evaluates with its own data.

The emitted JS calls helper functions, so the same expression runs on the workbook's figures
(for checking) or on the app's figures (live):
  R(col,row)            a Stock cell                     G(col,r1,r2) a Stock column range
  OB(col,date)          Order Book column, row for date  SUMIF(cc,r1,r2,crit,sc)
  MT(col,row)           a Monthly Targets cell           N(x)  IFERR(fn,alt)  COUNT(a)  MAX(...)  MIN(...)
  PD                    the plan date (Excel serial, Stock!$B$3)
Written 21 Nov 2026 for step 5 of the app (see docs/app-plan.md)."""
import re

TOK = re.compile(r"""\s*(?:
  (?P<str>"(?:[^"]|"")*")|
  (?P<ref>(?:'[^']+'|[A-Za-z_]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|
  (?P<num>\d+(?:\.\d+)?)|
  (?P<fn>[A-Z][A-Z0-9.]*)\(|
  (?P<op><>|<=|>=|[-+*/&=<>(),])
)""", re.X)


def tokens(f):
    f = f.lstrip("=")
    i, out = 0, []
    while i < len(f):
        m = TOK.match(f, i)
        if not m or m.end() == i:
            raise ValueError("cannot read formula at: " + f[i:i + 30])
        k = m.lastgroup
        out.append((k, m.group(k)))
        i = m.end()
    return out


class P:
    def __init__(self, f):
        self.t, self.i = tokens(f), 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self, v=None):
        tk = self.t[self.i]
        if v is not None and tk[1] != v:
            raise ValueError(f"expected {v}, got {tk}")
        self.i += 1
        return tk

    def expr(self):
        return self.cmp()

    def cmp(self):
        a = self.cat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            b = self.cat()
            a = f"({a}{ {'=': '==', '<>': '!='}.get(op, op) }{b})"
        return a

    def cat(self):
        a = self.add()
        while self.peek()[1] == "&":
            self.take()
            a = f"(String({a})+String({self.add()}))"
        return a

    def add(self):
        a = self.mul()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            a = f"({a}{op}{self.mul()})"
        return a

    def mul(self):
        a = self.una()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            a = f"({a}{op}{self.una()})"
        return a

    def una(self):
        if self.peek()[1] == "-":
            self.take()
            return f"(-{self.una()})"
        return self.prim()

    def args(self):
        out = []
        if self.peek()[1] == ")":
            self.take(); return out
        while True:
            out.append(self.expr())
            if self.take()[1] == ")": return out

    def prim(self):
        k, v = self.take()
        if k == "num": return v
        if k == "str": return '"' + v[1:-1].replace('""', '\\"') + '"'
        if k == "op" and v == "(":
            e = self.expr(); self.take(")"); return f"({e})"
        if k == "ref": return ref(v)
        if k == "fn": return call(v, self)
        raise ValueError(f"unexpected {k} {v}")


def split(v):
    sheet = None
    if "!" in v:
        sheet, v = v.split("!")
        sheet = sheet.strip("'")
    v = v.replace("$", "")
    if ":" in v:
        a, b = v.split(":")
        ca, ra = re.match(r"([A-Z]+)(\d+)", a).groups()
        cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
        return sheet, (ca, int(ra), cb, int(rb))
    c, r = re.match(r"([A-Z]+)(\d+)", v).groups()
    return sheet, (c, int(r))


def ref(v):
    sheet, p = split(v)
    if sheet in (None, "Stock"):
        if len(p) == 2:
            return "PD" if p == ("B", 3) else f'R("{p[0]}",{p[1]})'
        return f'G("{p[0]}",{p[1]},{p[3]})'
    if sheet == "Monthly Targets" and len(p) == 2:
        return f'MT("{p[0]}",{p[1]})'
    raise ValueError("reference not handled: " + v)


def call(fn, ps):
    if fn == "INDEX":                      # INDEX('Order Book'!$C$7:$C$156, MATCH(x, 'Order Book'!$A..., 0))
        k, rng = ps.take()
        sheet, p = split(rng)
        assert sheet == "Order Book", rng
        ps.take(",")
        k2, m = ps.take()
        assert m == "MATCH"
        date = ps.expr(); ps.take(","); ps.take(); ps.take(","); ps.expr(); ps.take(")"); ps.take(")")
        return f'OB("{p[0]}",{date})'
    if fn == "SUMIF":                      # SUMIF($T$13:$T$56, crit, $J$13:$J$56)
        _, r1 = ps.take(); ps.take(","); crit = ps.expr(); ps.take(","); _, r2 = ps.take(); ps.take(")")
        _, a = split(r1); _, b = split(r2)
        return f'SUMIF("{a[0]}",{a[1]},{a[3]},{crit},"{b[0]}")'
    a = ps.args()
    if fn == "IF": return f"({a[0]}?{a[1]}:{a[2] if len(a) > 2 else 'false'})"
    if fn == "IFERROR": return f"IFERR(()=>({a[0]}),{a[1]})"
    if fn in ("MAX", "MIN", "N", "COUNT", "SUM"): return f"{fn}({','.join(a)})"
    if fn == "SUMPRODUCT":                 # only SUMPRODUCT(INT(range/k)) occurs: the sum of whole units per row
        m = re.fullmatch(r'Math\.floor\(\((G\("[A-Z]+",\d+,\d+\))/(\d+(?:\.\d+)?)\)\)', a[0])
        if not m: raise ValueError("SUMPRODUCT form not handled: " + a[0])
        return f"SUMPI({m.group(1)},{m.group(2)})"
    if fn == "INT": return f"Math.floor({a[0]})"
    if fn == "ROUND": return f"(Math.round(({a[0]})*Math.pow(10,{a[1]}))/Math.pow(10,{a[1]}))"
    raise ValueError("function not handled: " + fn)


def to_js(formula):
    p = P(formula)
    e = p.expr()
    if p.i != len(p.t):
        raise ValueError("left over: " + str(p.t[p.i:]))
    return e


# the helper library, the same text the app and the checker run
HELPERS = r"""
function MAKE_CTX(cell, ob, mt, pd){
  const flat=a=>a.reduce((x,y)=>x.concat(Array.isArray(y)?flat(y):[y]),[]);
  const num=v=>typeof v==="number"&&isFinite(v);
  const C={PD:pd,
    R:(c,r)=>cell(c,r),
    G:(c,r1,r2)=>{ const o=[]; for(let r=r1;r<=r2;r++) o.push(cell(c,r)); return o; },
    OB:(c,d)=>ob(c,d), MT:(c,r)=>mt(c,r),
    N:v=>num(v)?v:0,
    IFERR:(f,alt)=>{ try{ const v=f(); return (typeof v==="number"&&!isFinite(v))?alt:v; }catch(e){ return alt; } },
    COUNT:(...a)=>flat(a).filter(num).length,
    MAX:(...a)=>{ const v=flat(a).filter(num); return v.length?Math.max(...v):0; },
    MIN:(...a)=>{ const v=flat(a).filter(num); return v.length?Math.min(...v):0; },
    SUM:(...a)=>flat(a).filter(num).reduce((x,y)=>x+y,0),
    SUMPI:(arr,k)=>arr.reduce((s,v)=>s+(num(v)?Math.floor(v/k):0),0),
    SUMIF:(cc,r1,r2,crit,sc)=>{ let s=0; for(let r=r1;r<=r2;r++){ if(cell(cc,r)==crit){ const v=cell(sc,r); if(num(v)) s+=v; } } return s; }};
  C.run=expr=>{ const f=new Function("PD","R","G","OB","MT","N","IFERR","COUNT","MAX","MIN","SUMIF","SUM","SUMPI","return "+expr);
    return f(C.PD,C.R,C.G,C.OB,C.MT,C.N,C.IFERR,C.COUNT,C.MAX,C.MIN,C.SUMIF,C.SUM,C.SUMPI); };
  return C;
}
"""
