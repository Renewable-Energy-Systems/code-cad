import os, time, logging
from logging.handlers import RotatingFileHandler
import pythoncom, win32com.client as wc
from win32com.client import VARIANT

# ── CONFIG ─────────────────────────────────────────────────────
CIRCLE_DIA   = 76.0
LABEL_THK    = 0.15
BAR_HEIGHT   = 3.0
GAP          = 6.0
DIM_DROP     = 2.0
TXT_H        = 2.5
TOL_SCALE    = 0.9
NOTE_H       = 2.5
ARROW_SZ     = 3.0
MTEXT_W      = 90
# horizontal nudges (in units of TXT_H)
DX_DIA = -4          # left = negative, right = positive
DX_THK = -7.5          # left = negative, right = positive
# colour / layer
COL_GEOM, COL_DIM, COL_CEN = 7, 6, 2
TEMPLATE     = "acad.dwt"
OUT_DWG      = "disc_76.dwg"
LOG_FILE     = "disc_76.log"
WAIT_S       = 30
# ───────────────────────────────────────────────────────────────

vt = lambda p: VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, p)

def make_logger(folder):
    lg = logging.getLogger("disc76"); lg.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s",
                            "%Y-%m-%d %H:%M:%S")
    for h in (logging.StreamHandler(),
              RotatingFileHandler(os.path.join(folder, LOG_FILE),
                                  maxBytes=2e5, backupCount=3, encoding="utf-8")):
        h.setFormatter(fmt); lg.addHandler(h)
    return lg

def wait_prop(obj, attr, lbl, log):
    end=time.time()+WAIT_S
    while True:
        try: return getattr(obj, attr)
        except (pythoncom.com_error, AttributeError):
            if time.time()>end: raise
            log.info(f"{lbl} not ready…"); pythoncom.PumpWaitingMessages(); time.sleep(.5)

def try_set(o, names, v, log):
    for n in names:
        if hasattr(o, n):
            try: setattr(o, n, v); return
            except pythoncom.com_error: pass
    log.warning(f"{o}: can’t set {names}")

def ensure_layer(layers, doc, name, colour, ltype, log):
    try: lyr=layers.Item(name)
    except pythoncom.com_error: lyr=layers.Add(name)
    try_set(lyr,["Color","ColorIndex"], colour, log)
    if ltype and ltype not in (lt.Name for lt in doc.Linetypes):
        try: doc.Linetypes.Load(ltype,"acad.lin")
        except pythoncom.com_error: pass
    try_set(lyr,["Linetype"], ltype, log); return lyr

# ────────────────────────── MAIN ───────────────────────────────
def main():
    root=os.path.dirname(os.path.abspath(__file__)); log=make_logger(root)
    pythoncom.CoInitialize(); acad=None
    try:
        acad=wc.DispatchEx("AutoCAD.Application"); try_set(acad,["Visible"],False,log)
        docs=wait_prop(acad,"Documents","Documents",log)
        try: doc=docs.Add(TEMPLATE)
        except pythoncom.com_error: doc=docs.Add()

        ms     = wait_prop(doc,"ModelSpace","ModelSpace",log)
        layers = wait_prop(doc,"Layers",    "Layers",    log)

        ensure_layer(layers,doc,"GEOM",COL_GEOM,"CONTINUOUS",log)
        ensure_layer(layers,doc,"CENTER",COL_CEN,"CENTER2",log)
        ensure_layer(layers,doc,"DIM",COL_DIM,"CONTINUOUS",log)

        r=CIRCLE_DIA/2
        ms.AddCircle(vt((0,0,0)), r).Layer="GEOM"
        top=-(r+GAP); bot=top-BAR_HEIGHT; L,R=-r,r
        for a,b in [((L,top),(R,top)),((L,bot),(R,bot)),
                    ((L,top),(L,bot)),((R,top),(R,bot))]:
            ms.AddLine(vt((*a,0)), vt((*b,0))).Layer="GEOM"

        ms.AddLine(vt((-1.1*r,0,0)),vt((1.1*r,0,0))).Layer="CENTER"
        ms.AddLine(vt((0,1.1*r,0)), vt((0,-1.1*r,0))).Layer="CENTER"

        dia = ms.AddDimAligned(vt((L,bot-DIM_DROP,0)),vt((R,bot-DIM_DROP,0)),
                               vt((0,bot-DIM_DROP-10,0)))
        thk = ms.AddDimAligned(vt((R,top,0)),vt((R,bot,0)),
                               vt((R+10,(top+bot)/2,0)))

        for d in (dia, thk):
            d.Layer="DIM"
            try_set(d,["TextHeight"],TXT_H,log)
            try_set(d,["ArrowheadSize","ArrowSize"],ARROW_SZ,log)
            try_set(d,["TextColor","Color"],COL_GEOM,log)

        try_set(dia,["PrimaryUnitsPrecision"],0,log)
        try_set(thk,["PrimaryUnitsPrecision"],2,log)
        try_set(thk,["TextRotation","Rotation"],1.570796,log)

        # ── tolerance MTEXTs ──
        tol_h = TXT_H*TOL_SCALE
        h_code=f"\\H{TOL_SCALE:.3f}x;"

        # diameter tolerance: center above value, then shift left DX_DIA
        d_pos = dia.TextPosition
        d_pt  = vt((d_pos[0] + DX_DIA*TXT_H, d_pos[1]+TXT_H*1.20, 0))
        m1 = ms.AddMText(d_pt, 30, f"{h_code}+0.0\\P-0.2")
        m1.AttachmentPoint = 5
        m1.Layer="DIM"; m1.Color=COL_GEOM
        try_set(m1,["TextHeight","Height"],tol_h,log)

        # thickness tolerance: center above 0.15, shift left DX_THK
        t_pos = thk.TextPosition
        t_pt  = vt((t_pos[0]+DX_THK*TXT_H, t_pos[1]+TXT_H*0.6, 0))
        m2 = ms.AddMText(t_pt, 30, f"{h_code}±0.05")
        m2.AttachmentPoint = 5
        m2.Rotation = 1.570796
        m2.Layer="DIM"; m2.Color=COL_GEOM
        try_set(m2,["TextHeight","Height"],tol_h,log)

        # note
        note=("VISUAL CRITERIA :\nBURRS AND EDGE CUTS ARE NOT ALLOWED.\n"
              "BLACK SPOTS ARE NOT ALLOWED\nDI ELECTRIC STRENGTH ________")
        note_m=ms.AddMText(vt((L-5,bot-25,0)),MTEXT_W,note)
        note_m.Layer="GEOM"; try_set(note_m,["TextHeight","Height"],NOTE_H,log)

        out=os.path.join(root,OUT_DWG); doc.SaveAs(out); doc.Close(False)
        log.info(f"✅  DWG saved → {out}")

    finally:
        if acad: acad.Quit()
        pythoncom.CoUninitialize()

if __name__=="__main__":
    main()
