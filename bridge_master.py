# ==============================================================================
# BRIDGE MASTER - THE ULTIMATE ENGINE (DXF MULTI-CASE + INTERACTIVE BUILDER)
# ==============================================================================
# This module strictly implements precise DXF parsing with full interactive
# manual overrides for every structural element, load, support, and section.
# ==============================================================================

import streamlit as st
import numpy as np
import pandas as pd
import io
import os
import re
import math
import tempfile
import time
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from PIL import Image, ImageChops
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from report_builder import insert_blue_banner, add_eq, append_pdf_stream_to_word
from config import SECTIONS_DB, STRUTS_DB

try:
    import ezdxf
except ImportError:
    st.error("⚠️ مكتبة 'ezdxf' غير موجودة! برجاء تثبيتها: pip install ezdxf")
    ezdxf = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

# =========================================================
# 0. Helper Functions & Styles
# =========================================================
def apply_plot_styles():
    matplotlib.rcParams['font.family'] = 'sans-serif'
    matplotlib.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans']
    matplotlib.rcParams['axes.linewidth'] = 0.3
    matplotlib.rcParams['font.size'] = 7

def get_short_name(sec_name):
    return re.sub(r'\s*\(.*?\)', '', sec_name).strip()

def crop_image_bbox(img_bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    bg = Image.new("RGBA", img.size, (255, 255, 255, 0))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if bbox: img = img.crop(bbox)
    out = io.BytesIO()
    img.save(out, format='PNG')
    return out.getvalue()

def safe_render_fig(fig):
    try:
        plt.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=400, bbox_inches='tight', pad_inches=0.0, transparent=True)
        return crop_image_bbox(buf.getvalue())
    finally:
        plt.close(fig)

def draw_reaction_arrow(ax, node_x, node_y, force_mag, axis_nx, axis_ny):
    if abs(force_mag) < 0.001: return
    arr_L = 0.5  
    sgn = np.sign(force_mag)
    dx, dy = sgn * axis_nx, sgn * axis_ny
    start_x, start_y = node_x - arr_L * dx, node_y - arr_L * dy
    arr_c = 'blue' if force_mag >= 0 else 'red'
    ax.arrow(start_x, start_y, arr_L * dx, arr_L * dy, length_includes_head=True, head_width=0.08, head_length=0.12, fc=arr_c, ec=arr_c, lw=0.8, zorder=5)
    ax.text(start_x - 0.15 * dx, start_y - 0.15 * dy, f"{force_mag:+.2f}", color=arr_c, fontsize=7, fontname='Arial', ha='center', va='center')

def eval_seg_point(seg, s_val):
    if seg.get('is_divided'):
        actual_s = s_val + seg.get('parent_offset', 0.0)
        return eval_seg_point(seg['parent_seg'], actual_s)

    L = seg.get('L', 0.0)
    s_val = min(max(s_val, 0.0), L)
    ratio = s_val / L if L > 1e-6 else 0.0
    
    if seg.get('is_dxf'):
        if seg.get('Shape Type') == 'Straight Line':
            p1, p2 = seg.get('abs_p1', (0,0)), seg.get('abs_p2', (0,0))
            px = p1[0] + ratio * (p2[0] - p1[0])
            py = p1[1] + ratio * (p2[1] - p1[1])
            th = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
            return px, py, th
        elif seg.get('Shape Type') == 'Curve (Arc & Radius)':
            c, r = seg['abs_c'], seg['abs_r']
            current_ang = seg['abs_sa'] + ratio * seg.get('sweep', 0)
            px = c[0] + r * math.cos(current_ang)
            py = c[1] + r * math.sin(current_ang)
            th = current_ang + math.pi / 2.0
            return px, py, th
    
    return 0.0, 0.0, 0.0

def get_closest_segment_exact(pt, segs):
    min_d, best_idx, best_s = 9999.0, 0, 0.0
    px, py = pt[0], pt[1]
    for idx, seg in enumerate(segs):
        if seg.get('is_divided'):
            temp_seg = seg['parent_seg']
            L_orig = temp_seg.get('L', 0.0)
        else:
            temp_seg = seg
            L_orig = temp_seg.get('L', 0.0)
            
        if 'abs_p1' in temp_seg:
            p1, p2 = np.array(temp_seg['abs_p1']), np.array(temp_seg['abs_p2'])
            v, w = p2 - p1, np.array([px, py]) - p1
            c2 = np.dot(v, v)
            ratio = max(0.0, min(1.0, np.dot(w, v) / c2 if c2 > 1e-6 else 0.0))
            proj = p1 + ratio * v
            d = np.linalg.norm(np.array([px, py]) - proj)
            if d < min_d:
                min_d, best_idx, best_s = d, idx, ratio * L_orig
                if seg.get('is_divided'): best_s -= seg.get('parent_offset', 0.0)
    return min_d, best_idx, best_s

# =========================================================
# 1. THE STRICT DXF PARSER (Cases, Areas, Loads)
# =========================================================
def parse_dxf_bridge_cases(file_bytes, loaded_width, conc_density):
    """
    استخراج هندسي دقيق لبيانات الكوبري من الـ DXF وتصنيفها إلى حالات (Cases).
    يفصل بين الحالات استناداً لخط الـ SEPARATOR.
    """
    if ezdxf is None: return None
    tmp_path = ""
    try:
        try: dxf_str = file_bytes.decode('utf-8')
        except: dxf_str = file_bytes.decode('cp1252', errors='ignore')
            
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf", mode='w', encoding='utf-8') as tmp:
            tmp.write(dxf_str)
            tmp_path = tmp.name
            
        doc = ezdxf.readfile(tmp_path)
        msp = doc.modelspace()
        
        layer_supp = 'SUPPORT'
        layer_text = 'TEXT_DATA'
        layer_sep = 'SEPARATOR'
        layer_frame = 'FRAME'
        layer_strut = 'PUSH_PULL'
        
        separators = []
        for e in msp:
            if e.dxftype() == 'LINE' and e.dxf.layer.upper() == layer_sep:
                separators.append((e.dxf.start.x + e.dxf.end.x) / 2000.0) 
                
        separators.sort()
        separators = [-999999.0] + separators + [999999.0]
        
        cases_raw = []
        for i in range(len(separators)-1):
            cases_raw.append({
                'min_x': separators[i], 'max_x': separators[i+1],
                'frames': [], 'struts': [], 'supports': [], 'cut_points': [], 's_texts': [], 'a_texts': []
            })
            
        for e in msp:
            layer = e.dxf.layer.upper()
            dxftype = e.dxftype()
            x_cad, y_cad = 0, 0
            is_valid_point = False
            
            # 1. قراءة النقاط الصريحة والبلوكات (Points & Inserts)
            if dxftype in ['POINT', 'CIRCLE']:
                x_cad = (e.dxf.location.x if dxftype=='POINT' else e.dxf.center.x) / 1000.0
                y_cad = (e.dxf.location.y if dxftype=='POINT' else e.dxf.center.y) / 1000.0
                is_valid_point = True
            elif dxftype == 'INSERT':
                x_cad, y_cad = e.dxf.insert.x / 1000.0, e.dxf.insert.y / 1000.0
                is_valid_point = True
                
            if is_valid_point:
                for c in cases_raw:
                    if c['min_x'] <= x_cad <= c['max_x']:
                        if layer == layer_supp: 
                            c['supports'].append({'x': x_cad, 'y': y_cad, 'type': 'Roller', 'angle': 0.0})
                        elif layer == layer_text: 
                            c['cut_points'].append({'x': x_cad, 'y': y_cad})
                        break

            # 2. قراءة النصوص لحساب المساحات والأطوال (Texts)
            elif dxftype in ['TEXT', 'MTEXT']:
                x_cad, y_cad = e.dxf.insert.x / 1000.0, e.dxf.insert.y / 1000.0
                for c in cases_raw:
                    if c['min_x'] <= x_cad <= c['max_x'] and layer == layer_text:
                        txt = e.text if dxftype == 'MTEXT' else e.dxf.text
                        txt = txt.upper().replace('\n', '').replace('\r', '')
                        s_m = re.search(r'S\s*(\d+)\s*=\s*([\d\.]+)', txt)
                        a_m = re.search(r'A\s*(\d+)\s*=\s*([\d\.]+)', txt)
                        # الوحدات تُقرأ كأمتار مباشرة دون قسمة على 1000
                        if s_m: c['s_texts'].append({'idx': int(s_m.group(1)), 'val': float(s_m.group(2)), 'x': x_cad, 'y': y_cad})
                        if a_m: c['a_texts'].append({'idx': int(a_m.group(1)), 'val': float(a_m.group(2)), 'x': x_cad, 'y': y_cad})
                        break
            
            # 3. تفجير الخطوط والمسارات (Lines & Polylines)
            elif dxftype in ['LINE', 'LWPOLYLINE', 'POLYLINE']:
                entities = list(e.virtual_entities()) if dxftype != 'LINE' else [e]
                for sub_e in entities:
                    if sub_e.dxftype() == 'LINE':
                        x1, y1 = sub_e.dxf.start.x / 1000.0, sub_e.dxf.start.y / 1000.0
                        x2, y2 = sub_e.dxf.end.x / 1000.0, sub_e.dxf.end.y / 1000.0
                        mid_x = (x1 + x2) / 2.0
                        for c in cases_raw:
                            if c['min_x'] <= mid_x <= c['max_x']:
                                if layer == layer_frame: c['frames'].append({'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2})
                                elif layer == layer_strut: c['struts'].append({'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2})
                                break

        # 4. ترجمة الحالات لمعلومات هندسية نهائية
        processed_cases = []
        for c_idx, c in enumerate(cases_raw):
            if not c['frames']: continue 
            
            base_segments = []
            for i, line in enumerate(c['frames']):
                L = math.hypot(line['x2'] - line['x1'], line['y2'] - line['y1'])
                base_segments.append({
                    'name': f"F{i+1}", 'master_idx': i, 'type': 'Straight Line', 'Shape Type': 'Straight Line',
                    'L': L, 'is_dxf': True, 'abs_p1': (line['x1'], line['y1']), 'abs_p2': (line['x2'], line['y2'])
                })
                
            loads = []
            # تسمية القطاعات وتوليد أحمالها العمودية للأسفل (Gravity)
            for s_txt in c['s_texts']:
                min_d, best_idx, _ = get_closest_segment_exact((s_txt['x'], s_txt['y']), base_segments)
                if min_d < 2.0: 
                    base_segments[best_idx]['name'] = f"S{s_txt['idx']}"
                    a_txt = next((a for a in c['a_texts'] if a['idx'] == s_txt['idx']), None)
                    if a_txt:
                        area = a_txt['val']
                        s_len = s_txt['val'] 
                        if s_len > 1e-4:
                            w_val = (area * conc_density * loaded_width) / s_len
                            loads.append({
                                'seg_idx': best_idx, 'category': 'Dead Load', 'type': 'Uniform',
                                'dir': 'Global Y (Vertical)', 'start': 0.0, 'end': base_segments[best_idx]['L'],
                                'w1': -abs(w_val), 'w2': -abs(w_val), 'target_mode': 'Single Segment'
                            })
                            
            struts_mapped = []
            for line in c['struts']:
                if line['y1'] > line['y2']: tx, ty, bx, by = line['x1'], line['y1'], line['x2'], line['y2']
                else: tx, ty, bx, by = line['x2'], line['y2'], line['x1'], line['y1']
                struts_mapped.append({'tx': tx, 'ty': ty, 'bx': bx, 'by': by, 'sec': list(STRUTS_DB.keys())[0] if STRUTS_DB else "PPH"})
                
            if c['supports']:
                c['supports'].sort(key=lambda sp: sp['x'])
                c['supports'][0]['type'] = 'Hinged'
                
            processed_cases.append({
                'title': f"Case {c_idx+1}",
                'segments': base_segments,
                'divided_segments': base_segments.copy(),
                'struts': struts_mapped,
                'supports': c['supports'],
                'cut_points': c['cut_points'],
                'loads': loads
            })
            
        return processed_cases
    except Exception as e:
        st.error(f"DXF Parsing Error: {e}")
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass

# =========================================================
# 2. Smart Topology & Division Engine
# =========================================================
def perform_smart_division(base_segments, supports, struts, cut_points=[]):
    """
    التقطيع البصري والتوبولوجي الدقيق عند نقاط الارتكاز للحصول على M, V دقيقين.
    """
    cut_points_dict = {i: {0.0, seg['L']} for i, seg in enumerate(base_segments)}
    for sp in supports:
        d_min, w_seg, w_s = get_closest_segment_exact((sp['x'], sp['y']), base_segments)
        if d_min < 0.30: cut_points_dict[w_seg].add(min(max(w_s, 0.0), base_segments[w_seg]['L']))
    for st in struts:
        dt, wt_seg, wt_s = get_closest_segment_exact((st['tx'], st['ty']), base_segments)
        if dt < 0.30: cut_points_dict[wt_seg].add(min(max(wt_s, 0.0), base_segments[wt_seg]['L']))
        db, wb_seg, wb_s = get_closest_segment_exact((st['bx'], st['by']), base_segments)
        if db < 0.30: cut_points_dict[wb_seg].add(min(max(wb_s, 0.0), base_segments[wb_seg]['L']))
    for cp in cut_points:
        d_min, w_seg, w_s = get_closest_segment_exact((cp['x'], cp['y']), base_segments)
        if d_min < 0.30: cut_points_dict[w_seg].add(min(max(w_s, 0.0), base_segments[w_seg]['L']))

    divided_segments = []
    sub_letters = "abcdefghijklmnopqrstuvwxyz"
    for m_idx, s_vals_set in sorted(cut_points_dict.items()):
        master_seg = base_segments[m_idx]
        sorted_s = sorted(list(s_vals_set))
        num_sub = len(sorted_s) - 1
        for k in range(num_sub):
            s_start, s_end = sorted_s[k], sorted_s[k+1]
            if s_end - s_start < 1e-4: continue
            sub_name = master_seg['name'] if num_sub == 1 else f"{master_seg['name']}-{sub_letters[k % 26]}"
            new_seg = master_seg.copy()
            new_seg.update({'name': sub_name, 'is_divided': True, 'parent_seg': master_seg, 'parent_offset': s_start, 'L': s_end - s_start, 'master_idx': m_idx})
            divided_segments.append(new_seg)
    return divided_segments

# =========================================================
# 3. Meshing & FEA Matrix Engine (True 2D Frame)
# =========================================================
def build_chain_mesh(segments, seg_sections, loads, struts, supports, cut_points=[], mesh_size=0.50):
    nodes = []
    elements = []
    nodal_loads = []
    node_tol = 0.01 
    
    def get_or_add_node(x, y):
        for i, n in enumerate(nodes):
            if abs(n[0] - x) < node_tol and abs(n[1] - y) < node_tol: return i
        nodes.append([x, y])
        return len(nodes) - 1

    support_injections = {i: [] for i in range(len(segments))}
    supports_list_out = []
    
    for sup in supports:
        sx, sy = sup['x'], sup['y']
        min_d, w_seg, w_s = get_closest_segment_exact((sx, sy), segments)
        if min_d < 0.30: support_injections[w_seg].append(w_s)
        nid = get_or_add_node(sx, sy)
        supports_list_out.append({'node': nid, 'type': sup.get('type', 'Roller'), 'angle': sup.get('angle', 0.0)})

    for cp in cut_points:
        min_d, w_seg, w_s = get_closest_segment_exact((cp['x'], cp['y']), segments)
        if min_d < 0.30: support_injections[w_seg].append(w_s)
        get_or_add_node(cp['x'], cp['y'])

    for st_idx, st_item in enumerate(struts):
        tx, ty, bx, by = st_item['tx'], st_item['ty'], st_item['bx'], st_item['by']
        dt, wt_seg, wt_s = get_closest_segment_exact((tx, ty), segments)
        if dt < 0.30: 
            support_injections[wt_seg].append(wt_s)
            tx, ty, _ = eval_seg_point(segments[wt_seg], wt_s)
            
        db, wb_seg, wb_s = get_closest_segment_exact((bx, by), segments)
        if db < 0.30: 
            support_injections[wb_seg].append(wb_s)
            bx, by, _ = eval_seg_point(segments[wb_seg], wb_s)
            
        top_node = get_or_add_node(tx, ty)
        bot_node = get_or_add_node(bx, by)
        elements.append({
            'type': 'truss', 'group': 'strut', 'sec': st_item.get('sec', 'Unknown'),
            'n1': bot_node, 'n2': top_node, 'E': 21000000.0, 'A': 0.001
        })

    for i, seg in enumerate(segments):
        L = seg['L']
        key_s_vals = [0.0, L] + support_injections[i]
        for ld in loads:
            if ld['seg_idx'] == i: key_s_vals.extend([ld['start'], ld['end']])
        for p in np.linspace(0, L, max(1, int(np.ceil(L / mesh_size)))+1): key_s_vals.append(p)
            
        keys = sorted(list(set([min(max(round(k, 4), 0.0), round(L, 4)) for k in key_s_vals])))
        node_indices = [get_or_add_node(*eval_seg_point(seg, s)[:2]) for s in keys]
        
        m_idx = seg.get('master_idx', i)
        sec_props = seg_sections[m_idx] if m_idx < len(seg_sections) else seg_sections[0]
        
        for j in range(len(keys)-1):
            n1, n2 = node_indices[j], node_indices[j+1]
            if n1 == n2: continue 
            s_mid = (keys[j] + keys[j+1]) / 2.0
            _, _, th_mid = eval_seg_point(seg, s_mid)
            c_t, s_t = np.cos(th_mid), np.sin(th_mid)
            p_x1, p_y1, p_x2, p_y2 = 0.0, 0.0, 0.0, 0.0
            
            for ld in loads:
                if ld['seg_idx'] == i and ld.get('type') != 'Point Load' and ld['start']-1e-4 <= s_mid <= ld['end']+1e-4:
                    L_ld = max(ld['end'] - ld['start'], 1e-5)
                    wa = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j] - ld['start']) / L_ld
                    wb = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j+1] - ld['start']) / L_ld
                    dir_str = ld.get('dir', 'Global Y (Vertical)').upper()
                    if 'Z' in dir_str or 'Y' in dir_str:
                        p_x1 += wa * s_t; p_y1 += wa * c_t
                        p_x2 += wb * s_t; p_y2 += wb * c_t
                    elif 'X' in dir_str:
                        p_x1 += wa * c_t; p_y1 -= wa * s_t
                        p_x2 += wb * c_t; p_y2 -= wb * s_t
                    else:
                        p_y1 += wa; p_y2 += wb
                        
            elements.append({
                'type': 'frame', 'group': 'segment', 'sec': sec_props['name'],
                'n1': n1, 'n2': n2, 'px1': p_x1, 'py1': p_y1, 'px2': p_x2, 'py2': p_y2,
                'E': sec_props['E'] * 10000.0, 'A': sec_props['A'], 'I': sec_props['I'] / 100000000.0,
                'seg_idx': i, 'L': keys[j+1] - keys[j]
            })
            
        for ld in loads:
            if ld.get('seg_idx') == i and ld.get('type') == 'Point Load':
                px, py, th_pt = eval_seg_point(seg, ld['start'])
                nid = get_or_add_node(px, py)
                dir_str = ld.get('dir', 'Global Y (Vertical)').upper()
                if 'Z' in dir_str or 'Y' in dir_str: nodal_loads.append({'node': nid, 'Fx': 0.0, 'Fy': ld['w1']})
                elif 'X' in dir_str: nodal_loads.append({'node': nid, 'Fx': ld['w1'], 'Fy': 0.0})
                else: 
                    c_pt, s_pt = np.cos(th_pt), np.sin(th_pt)
                    nodal_loads.append({'node': nid, 'Fx': -ld['w1']*s_pt, 'Fy': ld['w1']*c_pt})

    return nodes, elements, nodal_loads, supports_list_out

def solve_fea_engine(nodes, elements, nodal_loads, supports_list):
    NDOF = len(nodes) * 3
    K = np.zeros((NDOF, NDOF))
    F = np.zeros(NDOF)
    
    for el in elements:
        n1, n2 = el['n1'], el['n2']
        x1, y1 = nodes[n1][0], nodes[n1][1]
        x2, y2 = nodes[n2][0], nodes[n2][1]
        L = np.hypot(x2 - x1, y2 - y1)
        if L < 1e-5: continue
        c, s = (x2 - x1) / L, (y2 - y1) / L
        el['L'], el['c'], el['s'] = L, c, s
        E, A, I = el['E'], el['A'], el.get('I', 0.00005)
        
        T = np.array([[c, s, 0, 0, 0, 0], [-s, c, 0, 0, 0, 0], [0, 0, 1, 0, 0, 0], [0, 0, 0, c, s, 0], [0, 0, 0, -s, c, 0], [0, 0, 0, 0, 0, 1]])
        
        if el['type'] == 'truss':
            k_loc = np.zeros((6, 6))
            k_loc[0, 0] = E * A / L; k_loc[3, 3] = E * A / L
            k_loc[0, 3] = -E * A / L; k_loc[3, 0] = -E * A / L
        else:
            k_loc = np.array([
                [E*A/L, 0, 0, -E*A/L, 0, 0], 
                [0, 12*E*I/L**3, 6*E*I/L**2, 0, -12*E*I/L**3, 6*E*I/L**2],
                [0, 6*E*I/L**2, 4*E*I/L, 0, -6*E*I/L**2, 2*E*I/L], 
                [-E*A/L, 0, 0, E*A/L, 0, 0],
                [0, -12*E*I/L**3, -6*E*I/L**2, 0, 12*E*I/L**3, -6*E*I/L**2], 
                [0, 6*E*I/L**2, 2*E*I/L, 0, -6*E*I/L**2, 4*E*I/L]
            ])
            px1, py1, px2, py2 = el.get('px1',0), el.get('py1',0), el.get('px2',0), el.get('py2',0)
            f_loc = np.array([(2*px1+px2)*L/6, (7*py1+3*py2)*L/20, (3*py1+2*py2)*L**2/60, (px1+2*px2)*L/6, (3*py1+7*py2)*L/20, -(2*py1+3*py2)*L**2/60])
            f_glob = T.T @ f_loc
            dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
            for r in range(6): F[dof[r]] += f_glob[r]
            
        k_glob = T.T @ k_loc @ T
        dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
        for r in range(6):
            for col in range(6): K[dof[r], dof[col]] += k_glob[r, col]
                
    for nl in nodal_loads:
        F[3*nl['node']] += nl['Fx']
        F[3*nl['node']+1] += nl['Fy']

    K_orig = K.copy()
    fixed_dofs = []
    K_pen = 1e12
    for sup in supports_list:
        n, t, a = sup['node'], sup['type'], sup.get('angle', 0.0)
        if t == 'Fixed': fixed_dofs.extend([3*n, 3*n+1, 3*n+2])
        elif t == 'Hinged': fixed_dofs.extend([3*n, 3*n+1])
        elif t == 'Roller':
            rad = np.radians(a)
            nx, ny = -np.sin(rad), np.cos(rad) 
            K[3*n, 3*n] += K_pen * nx**2
            K[3*n+1, 3*n+1] += K_pen * ny**2
            K[3*n, 3*n+1] += K_pen * nx * ny
            K[3*n+1, 3*n] += K_pen * nx * ny

    free_dof = [i for i in range(NDOF) if i not in fixed_dofs]
    K_ff, F_f = K[np.ix_(free_dof, free_dof)], F[free_dof]
    
    U = np.zeros(NDOF)
    try: U[free_dof] = np.linalg.solve(K_ff, F_f)
    except: U[free_dof] = np.linalg.lstsq(K_ff, F_f, rcond=None)[0]
    
    R_reactions = K_orig @ U - F 
    
    for el in elements:
        if el.get('L', 0) < 1e-5: continue
        n1, n2 = el['n1'], el['n2']
        u_glob = U[[3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]]
        T = np.array([[el['c'], el['s'], 0, 0, 0, 0], [-el['s'], el['c'], 0, 0, 0, 0], [0, 0, 1, 0, 0, 0], [0, 0, 0, el['c'], el['s'], 0], [0, 0, 0, -el['s'], el['c'], 0], [0, 0, 0, 0, 0, 1]])
        u_loc = T @ u_glob
        el['internal'] = {'u_loc': u_loc}
        
        if el['type'] == 'truss':
            N_val = (el['E'] * el['A'] / el['L']) * (u_loc[3] - u_loc[0])
            xs = np.linspace(0, el['L'], 51)
            el['internal'].update({'N': np.full_like(xs, N_val), 'V': np.zeros_like(xs), 'M': np.zeros_like(xs), 'x': xs})
        else:
            k_loc = np.array([
                [el['E']*el['A']/el['L'],0,0,-el['E']*el['A']/el['L'],0,0], [0,12*el['E']*el.get('I')/el['L']**3,6*el['E']*el.get('I')/el['L']**2,0,-12*el['E']*el.get('I')/el['L']**3,6*el['E']*el.get('I')/el['L']**2], [0,6*el['E']*el.get('I')/el['L']**2,4*el['E']*el.get('I')/el['L'],0,-6*el['E']*el.get('I')/el['L']**2,2*el['E']*el.get('I')/el['L']], [-el['E']*el['A']/el['L'],0,0,el['E']*el['A']/el['L'],0,0], [0,-12*el['E']*el.get('I')/el['L']**3,-6*el['E']*el.get('I')/el['L']**2,0,12*el['E']*el.get('I')/el['L']**3,-6*el['E']*el.get('I')/el['L']**2], [0,6*el['E']*el.get('I')/el['L']**2,2*el['E']*el.get('I')/el['L'],0,-6*el['E']*el.get('I')/el['L']**2,4*el['E']*el.get('I')/el['L']]])
            px1, py1, px2, py2 = el.get('px1',0), el.get('py1',0), el.get('px2',0), el.get('py2',0)
            f_loc = np.array([(2*px1+px2)*el['L']/6, (7*py1+3*py2)*el['L']/20, (3*py1+2*py2)*el['L']**2/60, (px1+2*px2)*el['L']/6, (3*py1+7*py2)*el['L']/20, -(2*py1+3*py2)*el['L']**2/60])
            f_end = k_loc @ u_loc - f_loc
            
            xs = np.linspace(0, el['L'], 51)
            el['internal'].update({
                'N': -f_end[0] - (px1*xs + (px2-px1)*xs**2/(2*el['L'])), 
                'V': f_end[1] + (py1*xs + (py2-py1)*xs**2/(2*el['L'])), 
                'M': -f_end[2] + f_end[1]*xs + py1*xs**2/2.0 + (py2-py1)*xs**3/(6*el['L']), 
                'x': xs
            })
            
    return U, R_reactions, abs(np.sum(F[1::3]))
# =========================================================
# 2. Meshing & FEA Matrix Engine (True 2D Frame)
# =========================================================
def build_chain_mesh(segments, seg_sections, loads, struts, supports, cut_points=[], mesh_size=0.50):
    nodes = []
    elements = []
    nodal_loads = []
    node_tol = 0.01 
    
    def get_or_add_node(x, y):
        for i, n in enumerate(nodes):
            if abs(n[0] - x) < node_tol and abs(n[1] - y) < node_tol: return i
        nodes.append([x, y])
        return len(nodes) - 1

    support_injections = {i: [] for i in range(len(segments))}
    supports_list_out = []
    
    for sup in supports:
        sx, sy = sup['x'], sup['y']
        min_d, w_seg, w_s = get_closest_segment_exact((sx, sy), segments)
        if min_d < 0.30: support_injections[w_seg].append(w_s)
        nid = get_or_add_node(sx, sy)
        supports_list_out.append({'node': nid, 'type': sup.get('type', 'Roller'), 'angle': sup.get('angle', 0.0)})

    for cp in cut_points:
        min_d, w_seg, w_s = get_closest_segment_exact((cp['x'], cp['y']), segments)
        if min_d < 0.30: support_injections[w_seg].append(w_s)
        get_or_add_node(cp['x'], cp['y'])

    for st_idx, st_item in enumerate(struts):
        tx, ty, bx, by = st_item['tx'], st_item['ty'], st_item['bx'], st_item['by']
        dt, wt_seg, wt_s = get_closest_segment_exact((tx, ty), segments)
        if dt < 0.30: 
            support_injections[wt_seg].append(wt_s)
            tx, ty, _ = eval_seg_point(segments[wt_seg], wt_s)
            
        db, wb_seg, wb_s = get_closest_segment_exact((bx, by), segments)
        if db < 0.30: 
            support_injections[wb_seg].append(wb_s)
            bx, by, _ = eval_seg_point(segments[wb_seg], wb_s)
            
        top_node = get_or_add_node(tx, ty)
        bot_node = get_or_add_node(bx, by)
        elements.append({
            'type': 'truss', 'group': 'strut', 'sec': st_item.get('sec', 'Unknown'),
            'n1': bot_node, 'n2': top_node, 'E': 21000000.0, 'A': 0.001
        })

    for i, seg in enumerate(segments):
        L = seg['L']
        key_s_vals = [0.0, L] + support_injections[i]
        for ld in loads:
            if ld['seg_idx'] == i: key_s_vals.extend([ld['start'], ld['end']])
        for p in np.linspace(0, L, max(1, int(np.ceil(L / mesh_size)))+1): key_s_vals.append(p)
            
        keys = sorted(list(set([min(max(round(k, 4), 0.0), round(L, 4)) for k in key_s_vals])))
        node_indices = [get_or_add_node(*eval_seg_point(seg, s)[:2]) for s in keys]
        
        m_idx = seg.get('master_idx', i)
        sec_props = seg_sections[m_idx] if m_idx < len(seg_sections) else seg_sections[0]
        
        for j in range(len(keys)-1):
            n1, n2 = node_indices[j], node_indices[j+1]
            if n1 == n2: continue 
            s_mid = (keys[j] + keys[j+1]) / 2.0
            _, _, th_mid = eval_seg_point(seg, s_mid)
            c_t, s_t = np.cos(th_mid), np.sin(th_mid)
            p_x1, p_y1, p_x2, p_y2 = 0.0, 0.0, 0.0, 0.0
            
            for ld in loads:
                if ld['seg_idx'] == i and ld.get('type') != 'Point Load' and ld['start']-1e-4 <= s_mid <= ld['end']+1e-4:
                    L_ld = max(ld['end'] - ld['start'], 1e-5)
                    wa = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j] - ld['start']) / L_ld
                    wb = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j+1] - ld['start']) / L_ld
                    dir_str = ld.get('dir', 'Global Y (Vertical)').upper()
                    if 'Z' in dir_str or 'Y' in dir_str:
                        p_x1 += wa * s_t; p_y1 += wa * c_t
                        p_x2 += wb * s_t; p_y2 += wb * c_t
                    elif 'X' in dir_str:
                        p_x1 += wa * c_t; p_y1 -= wa * s_t
                        p_x2 += wb * c_t; p_y2 -= wb * s_t
                    else:
                        p_y1 += wa; p_y2 += wb
                        
            elements.append({
                'type': 'frame', 'group': 'segment', 'sec': sec_props['name'],
                'n1': n1, 'n2': n2, 'px1': p_x1, 'py1': p_y1, 'px2': p_x2, 'py2': p_y2,
                'E': sec_props['E'] * 10000.0, 'A': sec_props['A'], 'I': sec_props['I'] / 100000000.0,
                'seg_idx': i, 'L': keys[j+1] - keys[j]
            })
            
        for ld in loads:
            if ld.get('seg_idx') == i and ld.get('type') == 'Point Load':
                px, py, th_pt = eval_seg_point(seg, ld['start'])
                nid = get_or_add_node(px, py)
                dir_str = ld.get('dir', 'Global Y (Vertical)').upper()
                if 'Z' in dir_str or 'Y' in dir_str: nodal_loads.append({'node': nid, 'Fx': 0.0, 'Fy': ld['w1']})
                elif 'X' in dir_str: nodal_loads.append({'node': nid, 'Fx': ld['w1'], 'Fy': 0.0})
                else: 
                    c_pt, s_pt = np.cos(th_pt), np.sin(th_pt)
                    nodal_loads.append({'node': nid, 'Fx': -ld['w1']*s_pt, 'Fy': ld['w1']*c_pt})

    return nodes, elements, nodal_loads, supports_list_out

def solve_fea_engine(nodes, elements, nodal_loads, supports_list):
    NDOF = len(nodes) * 3
    K = np.zeros((NDOF, NDOF))
    F = np.zeros(NDOF)
    
    for el in elements:
        n1, n2 = el['n1'], el['n2']
        x1, y1 = nodes[n1][0], nodes[n1][1]
        x2, y2 = nodes[n2][0], nodes[n2][1]
        L = np.hypot(x2 - x1, y2 - y1)
        if L < 1e-5: continue
        c, s = (x2 - x1) / L, (y2 - y1) / L
        el['L'], el['c'], el['s'] = L, c, s
        E, A, I = el['E'], el['A'], el.get('I', 0.00005)
        
        T = np.array([[c, s, 0, 0, 0, 0], [-s, c, 0, 0, 0, 0], [0, 0, 1, 0, 0, 0], [0, 0, 0, c, s, 0], [0, 0, 0, -s, c, 0], [0, 0, 0, 0, 0, 1]])
        
        if el['type'] == 'truss':
            k_loc = np.zeros((6, 6))
            k_loc[0, 0] = E * A / L; k_loc[3, 3] = E * A / L
            k_loc[0, 3] = -E * A / L; k_loc[3, 0] = -E * A / L
        else:
            k_loc = np.array([
                [E*A/L, 0, 0, -E*A/L, 0, 0], 
                [0, 12*E*I/L**3, 6*E*I/L**2, 0, -12*E*I/L**3, 6*E*I/L**2],
                [0, 6*E*I/L**2, 4*E*I/L, 0, -6*E*I/L**2, 2*E*I/L], 
                [-E*A/L, 0, 0, E*A/L, 0, 0],
                [0, -12*E*I/L**3, -6*E*I/L**2, 0, 12*E*I/L**3, -6*E*I/L**2], 
                [0, 6*E*I/L**2, 2*E*I/L, 0, -6*E*I/L**2, 4*E*I/L]
            ])
            px1, py1, px2, py2 = el.get('px1',0), el.get('py1',0), el.get('px2',0), el.get('py2',0)
            f_loc = np.array([(2*px1+px2)*L/6, (7*py1+3*py2)*L/20, (3*py1+2*py2)*L**2/60, (px1+2*px2)*L/6, (3*py1+7*py2)*L/20, -(2*py1+3*py2)*L**2/60])
            f_glob = T.T @ f_loc
            dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
            for r in range(6): F[dof[r]] += f_glob[r]
            
        k_glob = T.T @ k_loc @ T
        dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
        for r in range(6):
            for col in range(6): K[dof[r], dof[col]] += k_glob[r, col]
                
    for nl in nodal_loads:
        F[3*nl['node']] += nl['Fx']
        F[3*nl['node']+1] += nl['Fy']

    K_orig = K.copy()
    fixed_dofs = []
    K_pen = 1e12
    for sup in supports_list:
        n, t, a = sup['node'], sup['type'], sup.get('angle', 0.0)
        if t == 'Fixed': fixed_dofs.extend([3*n, 3*n+1, 3*n+2])
        elif t == 'Hinged': fixed_dofs.extend([3*n, 3*n+1])
        elif t == 'Roller':
            rad = np.radians(a)
            nx, ny = -np.sin(rad), np.cos(rad) 
            K[3*n, 3*n] += K_pen * nx**2
            K[3*n+1, 3*n+1] += K_pen * ny**2
            K[3*n, 3*n+1] += K_pen * nx * ny
            K[3*n+1, 3*n] += K_pen * nx * ny

    free_dof = [i for i in range(NDOF) if i not in fixed_dofs]
    K_ff, F_f = K[np.ix_(free_dof, free_dof)], F[free_dof]
    
    U = np.zeros(NDOF)
    try: U[free_dof] = np.linalg.solve(K_ff, F_f)
    except: U[free_dof] = np.linalg.lstsq(K_ff, F_f, rcond=None)[0]
    
    R_reactions = K_orig @ U - F 
    
    for el in elements:
        if el.get('L', 0) < 1e-5: continue
        n1, n2 = el['n1'], el['n2']
        u_glob = U[[3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]]
        T = np.array([[el['c'], el['s'], 0, 0, 0, 0], [-el['s'], el['c'], 0, 0, 0, 0], [0, 0, 1, 0, 0, 0], [0, 0, 0, el['c'], el['s'], 0], [0, 0, 0, -el['s'], el['c'], 0], [0, 0, 0, 0, 0, 1]])
        u_loc = T @ u_glob
        el['internal'] = {'u_loc': u_loc}
        
        if el['type'] == 'truss':
            N_val = (el['E'] * el['A'] / el['L']) * (u_loc[3] - u_loc[0])
            xs = np.linspace(0, el['L'], 51)
            el['internal'].update({'N': np.full_like(xs, N_val), 'V': np.zeros_like(xs), 'M': np.zeros_like(xs), 'x': xs})
        else:
            k_loc = np.array([
                [el['E']*el['A']/el['L'],0,0,-el['E']*el['A']/el['L'],0,0], [0,12*el['E']*el.get('I')/el['L']**3,6*el['E']*el.get('I')/el['L']**2,0,-12*el['E']*el.get('I')/el['L']**3,6*el['E']*el.get('I')/el['L']**2], [0,6*el['E']*el.get('I')/el['L']**2,4*el['E']*el.get('I')/el['L'],0,-6*el['E']*el.get('I')/el['L']**2,2*el['E']*el.get('I')/el['L']], [-el['E']*el['A']/el['L'],0,0,el['E']*el['A']/el['L'],0,0], [0,-12*el['E']*el.get('I')/el['L']**3,-6*el['E']*el.get('I')/el['L']**2,0,12*el['E']*el.get('I')/el['L']**3,-6*el['E']*el.get('I')/el['L']**2], [0,6*el['E']*el.get('I')/el['L']**2,2*el['E']*el.get('I')/el['L'],0,-6*el['E']*el.get('I')/el['L']**2,4*el['E']*el.get('I')/el['L']]])
            px1, py1, px2, py2 = el.get('px1',0), el.get('py1',0), el.get('px2',0), el.get('py2',0)
            f_loc = np.array([(2*px1+px2)*el['L']/6, (7*py1+3*py2)*el['L']/20, (3*py1+2*py2)*el['L']**2/60, (px1+2*px2)*el['L']/6, (3*py1+7*py2)*el['L']/20, -(2*py1+3*py2)*el['L']**2/60])
            f_end = k_loc @ u_loc - f_loc
            
            xs = np.linspace(0, el['L'], 51)
            el['internal'].update({
                'N': -f_end[0] - (px1*xs + (px2-px1)*xs**2/(2*el['L'])), 
                'V': f_end[1] + (py1*xs + (py2-py1)*xs**2/(2*el['L'])), 
                'M': -f_end[2] + f_end[1]*xs + py1*xs**2/2.0 + (py2-py1)*xs**3/(6*el['L']), 
                'x': xs
            })
            
    # 💡 تم حل خطأ Unpacking بإرجاع القيمتين فقط اللتين تطلبهما دالة الواجهة
    return U, R_reactions
# =========================================================
# 3. Plotting & Report Generator
# =========================================================
def draw_base_geometry(ax, nodes, elements, supports_list, segments, show_names=False):
    for el in elements:
        n1, n2 = nodes[el['n1']], nodes[el['n2']]
        if el['type'] == 'truss':
            ax.plot([n1[0], n2[0]], [n1[1], n2[1]], color='red', linestyle='-', linewidth=0.8, zorder=1)
        else:
            ax.plot([n1[0], n2[0]], [n1[1], n2[1]], color='royalblue', linestyle='-', linewidth=1.5, zorder=1)
            
    for i, sup in enumerate(supports_list):
        x, y = nodes[sup['node']][0], nodes[sup['node']][1]
        t, ang_rad = sup['type'], math.radians(sup.get('angle', 0.0))
        c_a, s_a = math.cos(ang_rad), math.sin(ang_rad)
        
        def rot(px, py): return x+(px-x)*c_a-(py-y)*s_a, y+(px-x)*s_a+(py-y)*c_a
        
        if t == 'Fixed':
            ax.add_patch(Polygon([rot(x-0.1, y-0.1), rot(x+0.1, y-0.1), rot(x+0.1, y+0.1), rot(x-0.1, y+0.1)], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            ax.plot([rot(x-0.1, y)[0], rot(x+0.1, y)[0]], [rot(x-0.1, y)[1], rot(x+0.1, y)[1]], color='limegreen', lw=1.0, zorder=4)
        elif t == 'Hinged':
            ax.add_patch(Polygon([rot(x, y), rot(x+0.12, y-0.15), rot(x-0.12, y-0.15)], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            ax.plot([rot(x-0.17, y-0.15)[0], rot(x+0.17, y-0.15)[0]], [rot(x-0.17, y-0.15)[1], rot(x+0.17, y-0.15)[1]], color='limegreen', lw=1.0, zorder=4)
        elif t == 'Roller':
            ax.add_patch(Polygon([rot(x, y), rot(x+0.12, y-0.15), rot(x-0.12, y-0.15)], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            ax.add_patch(plt.Circle(rot(x, y-0.19), 0.04, facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            ax.plot([rot(x-0.17, y-0.23)[0], rot(x+0.17, y-0.23)[0]], [rot(x-0.17, y-0.23)[1], rot(x+0.17, y-0.23)[1]], color='limegreen', lw=1.0, zorder=4)

    if show_names and segments:
        for i, seg in enumerate(segments):
            mx, my, mth = eval_seg_point(seg, seg.get('L', 0)/2.0)
            rot_deg = math.degrees(mth)
            if rot_deg > 90: rot_deg -= 180
            elif rot_deg < -90: rot_deg += 180
            ax.text(mx - math.sin(mth)*0.3, my + math.cos(mth)*0.3, seg.get('name', f"S{i+1}"), color='dimgray', fontsize=6, ha='center', va='center', rotation=rot_deg, fontname='Arial')

# 💡 الدالة المسئولة عن الـ Live Preview قبل عمل الـ Analysis
def get_live_preview_image(nodes, elements, supports_list, loads, segments):
    apply_plot_styles()
    fig_ld, ax_ld = plt.subplots(figsize=(7, 4.5))
    ax_ld.set_aspect('equal', adjustable='datalim'); ax_ld.axis('off')
    draw_base_geometry(ax_ld, nodes, elements, supports_list, segments, show_names=True)
    
    for ld in loads:
        i = ld.get('seg_idx', 0)
        if i >= len(segments): continue
        w1, w2 = ld.get('w1', 0.0), ld.get('w2', 0.0)
        s_vals = np.linspace(ld.get('start', 0), ld.get('end', 0), 15)
        poly_pts, top_pts = [], []
        for sv in s_vals:
            px, py, th = eval_seg_point(segments[i], sv)
            L_load = max(1e-5, (ld.get('end', 0) - ld.get('start', 0)))
            w_val = (w1 + (w2 - w1) * (sv - ld.get('start', 0)) / L_load) * 0.05
            poly_pts.append((px, py))
            
            dir_str = ld.get('dir', 'Global Y (Vertical)').upper()
            if 'Y' in dir_str or 'Z' in dir_str:
                top_pts.append((px, py + w_val)) 
            elif 'X' in dir_str:
                top_pts.append((px + w_val, py))
            else:
                c, s = math.cos(th), math.sin(th)
                top_pts.append((px + s * w_val, py - c * w_val))
                
        poly_pts.extend(top_pts[::-1])
        if len(poly_pts) > 2:
            ax_ld.add_patch(Polygon(poly_pts, facecolor='blue', edgecolor='blue', alpha=0.15, lw=0.8, zorder=2))
            
    return safe_render_fig(fig_ld)

def plot_sap2000_diagrams(nodes, elements, R_reactions, scales, supports_list, loads, segments):
    apply_plot_styles()
    figs_dict = {}
    
    # 1. Loads Diagram
    figs_dict['L'] = get_live_preview_image(nodes, elements, supports_list, loads, segments)
    
    # 2. Reactions
    fig_r, ax_r = plt.subplots(figsize=(7, 4.5))
    ax_r.set_aspect('equal', adjustable='datalim'); ax_r.axis('off')
    draw_base_geometry(ax_r, nodes, elements, supports_list, segments)
    for sup in supports_list:
        n, ang = sup['node'], math.radians(sup.get('angle', 0.0))
        Rx, Ry = R_reactions[3*n], R_reactions[3*n+1]
        x, y = nodes[n][0], nodes[n][1]
        R_loc_x, R_loc_y = Rx * math.cos(ang) + Ry * math.sin(ang), -Rx * math.sin(ang) + Ry * math.cos(ang)
        if sup['type'] == 'Roller': draw_reaction_arrow(ax_r, x, y, R_loc_y, -math.sin(ang), math.cos(ang))
        else:
            draw_reaction_arrow(ax_r, x, y, R_loc_x, math.cos(ang), math.sin(ang))
            draw_reaction_arrow(ax_r, x, y, R_loc_y, -math.sin(ang), math.cos(ang))
            
    figs_dict['R'] = safe_render_fig(fig_r)
    
    # 3. Forces
    def create_force_plot(val_key, scale, c_pos, c_neg):
        fig_f, ax_f = plt.subplots(figsize=(7, 4.5))
        ax_f.set_aspect('equal', adjustable='datalim'); ax_f.axis('off')
        draw_base_geometry(ax_f, nodes, elements, supports_list, segments)
        for el in elements:
            n1, n2 = el['n1'], el['n2'] 
            x1, y1 = nodes[n1][0], nodes[n1][1]
            x2, y2 = nodes[n2][0], nodes[n2][1]
            c, s = el.get('c', 1.0), el.get('s', 0.0)
            xs = el.get('internal', {}).get('x', [])
            vals = el.get('internal', {}).get(val_key, [])
            
            if len(vals) == 0 or np.all(np.abs(vals) < 1e-6): continue
            plot_vals = -vals if val_key != 'N' else vals
            px, py = x1 + c*xs - s*plot_vals*scale, y1 + s*xs + c*plot_vals*scale
            
            for k in range(len(px)-1):
                ax_f.plot([px[k], px[k+1]], [py[k], py[k+1]], color=c_pos if vals[k] >= 0 else c_neg, lw=0.8)
            ax_f.plot([x1, px[0]], [y1, py[0]], color=c_pos if vals[0]>=0 else c_neg, lw=0.8)
            ax_f.plot([x2, px[-1]], [y2, py[-1]], color=c_pos if vals[-1]>=0 else c_neg, lw=0.8)
            
            mid = len(vals)//2
            if abs(vals[mid]) > 0.1: 
                ax_f.text(px[mid], py[mid], f"{vals[mid]:+.2f}", fontsize=6, color=c_pos if vals[mid]>=0 else c_neg, ha='center', va='center')
                
        return safe_render_fig(fig_f)

    figs_dict['N'] = create_force_plot('N', scales['N'], 'blue', 'red')
    figs_dict['V'] = create_force_plot('V', scales['V'], 'blue', 'red')
    figs_dict['M'] = create_force_plot('M', scales['M'], 'blue', 'red')
    
    return figs_dict

def generate_multi_case_report(cases_data, proj_info):
    doc = Document("Acrow_Template.docx") if os.path.exists("Acrow_Template.docx") else Document()
    def force_ltr_left(p):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '0')
        pPr.append(bidi)

    def add_line(text, bold=False, size=11):
        p = doc.add_paragraph()
        force_ltr_left(p)
        r = p.add_run(text)
        r.font.name, r.font.size, r.font.bold, r.font.rtl = 'Arial', Pt(size), bold, False
        
    add_line("BRIDGE FORMWORK MULTI-CASE ANALYSIS", bold=True, size=16)
    
    for case in cases_data:
        doc.add_page_break()
        add_line(f"ANALYSIS DIAGRAMS FOR {case['title'].upper()}", bold=True, size=14)
        doc.add_paragraph()
        
        for key, label in [('L', "APPLIED LOADS & GEOMETRY"), ('N', "AXIAL FORCE DIAGRAM (kN)"), ('M', "BENDING MOMENT DIAGRAM (kN.m)"), ('V', "SHEAR FORCE DIAGRAM (kN)"), ('R', "SUPPORT REACTIONS (kN)")]:
            if key in case['img_bufs']:
                add_line(label, bold=True, size=12)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(io.BytesIO(case['img_bufs'][key]), width=Cm(16.5))
                doc.add_paragraph()
                
    out = io.BytesIO()
    doc.save(out)
    return out

# =========================================================
# 4. Main UI Module (Dynamic Extractor & Interactive Editor)
# =========================================================
def render_bridge_module(proj_info):
    st.markdown("## 🌉 Bridge Formwork (True 2D DXF + Interactive Editor)")
    
    mode = st.radio("Select Input Mode:", ["1. Multi-Case DXF Auto-Extractor 🪄", "2. Single-Case Manual Builder 🛠️"], horizontal=True)
    st.markdown("---")

    if "DXF" in mode:
        st.info("💡 **Smart Engine:** Upload DXF. The AI extracts everything (Frames, Struts, Supports, Auto-Loads). You can then **fully edit** any support, section, or load before running the FEA!")
        
        c1, c2 = st.columns(2)
        loaded_width = c1.number_input("Loaded Width (m) for Load Calculation", value=1.30, step=0.05)
        conc_density = c2.number_input("Concrete Density (kN/m³)", value=25.0, step=0.5)
        
        uploaded_dxf = st.file_uploader("📥 Upload DXF File (.dxf) with ACROW Layers", type=['dxf'])
        
        if uploaded_dxf:
            if st.button("🚀 Process DXF & Extract Data", type="primary", use_container_width=True):
                with st.spinner("Parsing DXF true 2D geometry & Computing Automatic Loads..."):
                    cases_data = parse_dxf_bridge_cases(uploaded_dxf.getvalue(), loaded_width, conc_density)
                    
                if cases_data:
                    st.session_state.bridge_cases = cases_data
                    st.success(f"✅ Successfully extracted {len(cases_data)} structural 2D case(s) from the DXF!")
                    st.rerun()
                else:
                    st.error("❌ Failed to parse DXF. Please ensure layers (FRAME, PUSH_PULL, SUPPORT, TEXT_DATA, SEPARATOR) are correct.")

        if 'bridge_cases' in st.session_state:
            st.markdown("### 🎛️ Customize Global Diagram Scales")
            c_s1, c_s2, c_s3 = st.columns(3)
            sc_n = c_s1.slider("Axial Scale", 0.001, 0.050, 0.010, step=0.001)
            sc_v = c_s2.slider("Shear Scale", 0.001, 0.050, 0.010, step=0.001)
            sc_m = c_s3.slider("Moment Scale", 0.001, 0.100, 0.010, step=0.001)
            
            global_sec = {'name': "Soldier U100", 'E': 2100.0, 'A': 34.3 / 10000.0, 'I': 412.0 / 100000000.0, 'Mall': 13.1, 'Qall': 100.8}
            
            tabs = st.tabs([c['title'] for c in st.session_state.bridge_cases])
            all_cases_ready = []
            
            for c_idx, tab in enumerate(tabs):
                case = st.session_state.bridge_cases[c_idx]
                
                with tab:
                    st.markdown(f"#### 🛠️ Interactive Editor for {case['title']}")
                    c_edit, c_view = st.columns([1.2, 1.8])
                    
                    with c_edit:
                        with st.expander(f"🔗 Edit Supports ({len(case['supports'])})", expanded=False):
                            for i, sup in enumerate(case['supports']):
                                s1, s2, s3, s4 = st.columns([1,1,1.2,1])
                                sup['x'] = s1.number_input("X(m)", value=float(sup['x']), step=0.1, key=f"sx_{c_idx}_{i}")
                                sup['y'] = s2.number_input("Y(m)", value=float(sup['y']), step=0.1, key=f"sy_{c_idx}_{i}")
                                t_opts = ["Hinged", "Roller", "Fixed"]
                                sup['type'] = s3.selectbox("Type", t_opts, index=t_opts.index(sup['type']) if sup['type'] in t_opts else 1, key=f"st_{c_idx}_{i}")
                                sup['angle'] = s4.number_input("Ang(°)", value=float(sup.get('angle',0.0)), step=15.0, key=f"sa_{c_idx}_{i}")

                        if 'sec_overrides' not in case:
                            case['sec_overrides'] = [global_sec.copy() for _ in range(len(case['segments']))]
                            
                        with st.expander("📏 Edit Sections", expanded=False):
                            st.info("Default section is Soldier U100. Override below if needed.")
                            seg_names = [s['name'] for s in case['segments']]
                            override_segs = st.multiselect("Select segments to override:", seg_names, key=f"ovr_seg_{c_idx}")
                            if override_segs:
                                sec_choices = ["Custom Section", "Acrow Beam S12"]
                                sel_sec = st.radio("Override Profile:", sec_choices, key=f"ovr_rad_{c_idx}")
                                if sel_sec == "Custom Section":
                                    o1, o2, o3, o4 = st.columns(4)
                                    o_A = o1.number_input("A (cm2)", value=50.0, key=f"oa_{c_idx}")
                                    o_I = o2.number_input("I (cm4)", value=1200.0, key=f"oi_{c_idx}")
                                    o_M = o3.number_input("Mall", value=30.0, key=f"om_{c_idx}")
                                    o_Q = o4.number_input("Qall", value=150.0, key=f"oq_{c_idx}")
                                    o_sec = {'name': "Custom", 'E': 2100.0, 'A': o_A/10000.0, 'I': o_I/100000000.0, 'Mall': o_M, 'Qall': o_Q}
                                else:
                                    o_sec = {'name': "S12", 'E': 2100.0, 'A': 20.0/10000.0, 'I': 800.0/100000000.0, 'Mall': 15.0, 'Qall': 80.0}
                                    
                                for s_name in override_segs:
                                    idx_s = seg_names.index(s_name)
                                    case['sec_overrides'][idx_s] = o_sec.copy()

                        with st.expander(f"📐 Edit Struts ({len(case['struts'])})", expanded=False):
                            strut_opts = list(STRUTS_DB.keys()) if STRUTS_DB else ["PPH"]
                            for i, stt in enumerate(case['struts']):
                                s1, s2, s3, s4, s5 = st.columns([1,1,1,1,1.5])
                                stt['tx'] = s1.number_input("TX", value=float(stt['tx']), step=0.1, key=f"ttx_{c_idx}_{i}")
                                stt['ty'] = s2.number_input("TY", value=float(stt['ty']), step=0.1, key=f"tty_{c_idx}_{i}")
                                stt['bx'] = s3.number_input("BX", value=float(stt['bx']), step=0.1, key=f"tbx_{c_idx}_{i}")
                                stt['by'] = s4.number_input("BY", value=float(stt['by']), step=0.1, key=f"tby_{c_idx}_{i}")
                                stt['sec'] = s5.selectbox("Sec", strut_opts, index=0, key=f"tsec_{c_idx}_{i}")

                        with st.expander(f"⬇️ Edit Applied Loads ({len(case['loads'])})", expanded=True):
                            dir_opts = ["Global X (Horizontal)", "Global Y (Vertical)", "Global Z (Vertical)", "Local Z (Perpendicular)"]
                            seg_names = [s['name'] for s in case['segments']]
                            for i, ld in enumerate(case['loads']):
                                st.markdown(f"**Load {i+1}**")
                                l1, l2, l3, l4 = st.columns([1.5, 1.5, 1, 1])
                                idx_s = ld['seg_idx'] if ld['seg_idx'] < len(seg_names) else 0
                                s_name = l1.selectbox("Target Seg", seg_names, index=idx_s, key=f"lsg_{c_idx}_{i}")
                                ld['seg_idx'] = seg_names.index(s_name)
                                
                                idx_dir = dir_opts.index(ld.get('dir', 'Global Y (Vertical)')) if ld.get('dir') in dir_opts else 1
                                ld['dir'] = l2.selectbox("Direction", dir_opts, index=idx_dir, key=f"ldr_{c_idx}_{i}")
                                
                                ld['w1'] = l3.number_input("W1", value=float(ld['w1']), step=1.0, key=f"lw1_{c_idx}_{i}")
                                ld['w2'] = l4.number_input("W2", value=float(ld.get('w2', ld['w1'])), step=1.0, key=f"lw2_{c_idx}_{i}")
                                ld['type'] = 'Uniform'
                                
                            if st.button("➕ Add Manual Load Item", key=f"btn_add_ld_{c_idx}"):
                                case['loads'].append({'seg_idx': 0, 'category': 'Dead Load', 'type': 'Uniform', 'dir': 'Global Y (Vertical)', 'start': 0.0, 'end': case['segments'][0]['L'], 'w1': -10.0, 'w2': -10.0})
                                st.rerun()

                    with c_view:
                        # 💡 التحديث المنتظر: Live Geometry Preview Before FEA
                        st.markdown("<h4 style='text-align: center; color: #1e3d59;'>Live Geometry & Loads</h4>", unsafe_allow_html=True)
                        
                        prev_nodes, prev_elements, _, prev_supps = build_chain_mesh(
                            case['segments'], case['sec_overrides'], case['loads'], case['struts'], case['supports'], case.get('cut_points', [])
                        )
                        prev_img = get_live_preview_image(prev_nodes, prev_elements, prev_supps, case['loads'], case['segments'])
                        st.image(prev_img, use_container_width=True)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button(f"🚀 Run FEA Analysis for {case['title']}", type="primary", use_container_width=True, key=f"btn_run_{c_idx}"):
                            with st.spinner(f"Solving True 2D FEA for {case['title']}..."):
                                nodes, elements, nodal_loads, supports_list = build_chain_mesh(case['segments'], case['sec_overrides'], case['loads'], case['struts'], case['supports'], case.get('cut_points', []))
                                # تم حل مشكلة الـ ValueError هنا
                                U, R = solve_fea_engine(nodes, elements, nodal_loads, supports_list)[:2]
                                
                                img_bufs = plot_sap2000_diagrams(nodes, elements, R, {'N': sc_n, 'V': sc_v, 'M': sc_m}, supports_list, case['loads'], case['segments'])
                                case['img_bufs'] = img_bufs
                                
                                cc1, cc2 = st.columns(2)
                                cc1.image(img_bufs['M'], caption="Bending Moment (kN.m)")
                                cc2.image(img_bufs['V'], caption="Shear Force (kN)")
                                cc3, cc4 = st.columns(2)
                                cc3.image(img_bufs['N'], caption="Axial Force (kN)")
                                cc4.image(img_bufs['R'], caption="Support Reactions (kN)")
                                
                    all_cases_ready.append(case)
                            
            st.markdown("---")
            if st.button("📥 Download Multi-Case Word Report", type="primary", use_container_width=True):
                with st.spinner("Compiling Professional Calculation Sheet..."):
                    doc_out = generate_multi_case_report(all_cases_ready, proj_info)
                    st.download_button("💾 Save DXF Multi-Case Report", data=doc_out.getvalue(), file_name="Acrow_Bridge_DXF_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    # =========================================================
    # 💡 5. MANUAL BUILDER MODE (Full Interactive from Advanced)
    # =========================================================
    else:
        st.info("🛠️ **Interactive Manual Builder:** Define segments, supports, struts, and loads manually if you don't have a DXF file.")
        
        if 'man_segments' not in st.session_state:
            st.session_state.man_segments = [{'name': 'S1', 'L': 3.0, 'type': 'Straight Line'}]
        if 'man_supports' not in st.session_state:
            st.session_state.man_supports = [{'x': 0.0, 'y': 0.0, 'type': 'Hinged', 'angle': 0.0}, {'x': 3.0, 'y': 0.0, 'type': 'Roller', 'angle': 0.0}]
        if 'man_struts' not in st.session_state:
            st.session_state.man_struts = []
        if 'man_loads' not in st.session_state:
            st.session_state.man_loads = []

        c_in, c_plot = st.columns([1.2, 1.8])
        
        with c_in:
            st.markdown("### 1. Segments (Base Frames)")
            for i, seg in enumerate(st.session_state.man_segments):
                c_s1, c_s2 = st.columns([2, 1])
                seg['L'] = c_s1.number_input(f"{seg['name']} Length (m)", value=float(seg.get('L', 3.0)), step=0.1, key=f"ms_l_{i}")
                seg['master_idx'] = i
                seg['abs_p1'] = (sum([s['L'] for s in st.session_state.man_segments[:i]]), 0.0)
                seg['abs_p2'] = (seg['abs_p1'][0] + seg['L'], 0.0)
                seg['Shape Type'] = 'Straight Line'
            
            if st.button("➕ Add Segment"):
                st.session_state.man_segments.append({'name': f"S{len(st.session_state.man_segments)+1}", 'L': 3.0, 'type': 'Straight Line'})
                st.rerun()

            st.markdown("### 2. Supports")
            for i, sup in enumerate(st.session_state.man_supports):
                c_su1, c_su2, c_su3, c_su4 = st.columns([1, 1, 1.2, 0.8])
                sup['x'] = c_su1.number_input(f"J{i+1} X", value=float(sup.get('x', 0.0)), step=0.1, key=f"msup_x_{i}")
                sup['y'] = c_su2.number_input(f"J{i+1} Y", value=float(sup.get('y', 0.0)), step=0.1, key=f"msup_y_{i}")
                sup['type'] = c_su3.selectbox(f"J{i+1} Type", ["Hinged", "Roller", "Fixed"], index=0 if sup['type']=='Hinged' else 1, key=f"msup_t_{i}")
                if c_su4.button("❌ Del", key=f"del_msup_{i}"):
                    st.session_state.man_supports.pop(i); st.rerun()
            if st.button("➕ Add Support"):
                st.session_state.man_supports.append({'x': 0.0, 'y': 0.0, 'type': 'Roller', 'angle': 0.0})
                st.rerun()
                
            st.markdown("### 3. Struts (Push-Pulls)")
            strut_opts = list(STRUTS_DB.keys()) if STRUTS_DB else ["PPH"]
            for i, ds in enumerate(st.session_state.man_struts):
                c1, c2, c3, c4, c5 = st.columns([1.5, 1.5, 1.5, 1.5, 0.4])
                ds['tx'] = c1.number_input("Top X", value=float(ds.get('tx', 0.0)), step=0.1, key=f"mst_tx_{i}")
                ds['ty'] = c2.number_input("Top Y", value=float(ds.get('ty', 3.0)), step=0.1, key=f"mst_ty_{i}")
                ds['bx'] = c3.number_input("Bot X", value=float(ds.get('bx', 1.0)), step=0.1, key=f"mst_bx_{i}")
                ds['by'] = c4.number_input("Bot Y", value=float(ds.get('by', 0.0)), step=0.1, key=f"mst_by_{i}")
                ds['sec'] = "PPH"
                if c5.button("❌", key=f"del_mst_{i}"):
                    st.session_state.man_struts.pop(i); st.rerun()
            if st.button("➕ Add Strut"):
                st.session_state.man_struts.append({'tx': 0.0, 'ty': 3.0, 'bx': 1.0, 'by': 0.0, 'sec': 'PPH'})
                st.rerun()

            st.markdown("### 4. Applied Loads")
            dir_opts = ["Global X (Horizontal)", "Global Y (Vertical)", "Global Z (Vertical)", "Local Z (Perpendicular)"]
            for i, ld in enumerate(st.session_state.man_loads):
                c_l1, c_l2, c_l3, c_l4 = st.columns([1.2, 1.2, 1, 0.8])
                seg_names = [s['name'] for s in st.session_state.man_segments]
                target = c_l1.selectbox("Target Seg", seg_names, key=f"mld_s_{i}")
                ld['seg_idx'] = seg_names.index(target)
                ld['dir'] = c_l2.selectbox("Dir", dir_opts, index=1, key=f"mld_d_{i}")
                ld['w1'] = c_l3.number_input("W (kN/m)", value=float(ld.get('w1', -10.0)), step=1.0, key=f"mld_w_{i}")
                ld['w2'] = ld['w1']
                ld['start'] = 0.0
                ld['end'] = st.session_state.man_segments[ld['seg_idx']]['L']
                ld['type'] = 'Uniform'
                ld['category'] = 'Dead Load'
                if c_l4.button("❌ Del", key=f"del_mld_{i}"):
                    st.session_state.man_loads.pop(i); st.rerun()
            if st.button("➕ Add Load"):
                st.session_state.man_loads.append({'seg_idx': 0, 'w1': -10.0, 'dir': 'Global Y (Vertical)'})
                st.rerun()

        with c_plot:
            st.markdown("<h4 style='text-align: center; color: #1e3d59;'>Live Geometry & Loads</h4>", unsafe_allow_html=True)
            global_sec = {'name': "Soldier U100", 'E': 2100.0, 'A': 34.3 / 10000.0, 'I': 412.0 / 100000000.0, 'Mall': 13.1, 'Qall': 100.8}
            active_sections = [global_sec] * len(st.session_state.man_segments)
            
            # 💡 تحديث واجهة المانيوال لتظهر أيضاً كـ Live Preview قبل التحليل
            p_nodes, p_elements, _, p_supports_list = build_chain_mesh(
                st.session_state.man_segments, active_sections, st.session_state.man_loads, 
                st.session_state.man_struts, st.session_state.man_supports, []
            )
            man_prev_img = get_live_preview_image(p_nodes, p_elements, p_supports_list, st.session_state.man_loads, st.session_state.man_segments)
            st.image(man_prev_img, use_container_width=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 Run Advanced FEA", type="primary", use_container_width=True):
                with st.spinner("Solving FEA..."):
                    nodes, elements, nodal_loads, supports_list = build_chain_mesh(
                        st.session_state.man_segments, active_sections, st.session_state.man_loads, 
                        st.session_state.man_struts, st.session_state.man_supports, []
                    )
                    # تم حل الـ Unpacking Error هنا أيضاً
                    U, R = solve_fea_engine(nodes, elements, nodal_loads, supports_list)[:2]
                    img_bufs = plot_sap2000_diagrams(nodes, elements, R, {'N': 0.01, 'V': 0.01, 'M': 0.01}, supports_list, st.session_state.man_loads, st.session_state.man_segments)
                    st.session_state.man_img_bufs = img_bufs
                    st.session_state.man_case_data = [{'title': 'Manual Case', 'img_bufs': img_bufs}]
                    
            if 'man_img_bufs' in st.session_state:
                c_p1, c_p2 = st.columns(2)
                c_p1.image(st.session_state.man_img_bufs['M'], caption="Bending Moment")
                c_p2.image(st.session_state.man_img_bufs['V'], caption="Shear Force")
                c_p3, c_p4 = st.columns(2)
                c_p3.image(st.session_state.man_img_bufs['N'], caption="Axial Force")
                c_p4.image(st.session_state.man_img_bufs['R'], caption="Support Reactions")
                
                st.markdown("---")
                if st.button("📥 Download Word Report", type="primary", use_container_width=True):
                    doc_out = generate_multi_case_report(st.session_state.man_case_data, proj_info)
                    st.download_button("💾 Save Manual Case Report", data=doc_out.getvalue(), file_name="Acrow_Bridge_Manual_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)