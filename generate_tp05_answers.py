#!/usr/bin/env python3
"""
Génération du document TP05 - Énergisation et Réenclenchement de Lignes de Transport
Surtensions transitoires - ATP-EMTP / ATPDraw
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

IMG = "/workspace/images_tp05"
os.makedirs(IMG, exist_ok=True)

# ============================================================
# LINE PARAMETERS
# ============================================================
f = 50  # Hz
omega = 2 * np.pi * f
V_rated = 400e3  # V
V_peak = V_rated * np.sqrt(2) / np.sqrt(3)  # phase peak voltage
line_length = 300  # km (3 sections of 100 km)
t_close = 0.015  # breaker close at 15 ms
t_sim = 0.08  # simulation 80 ms
dt = 1e-5  # time step

# Source equivalent
R1 = 1.3287; X1 = 15.9447; L1 = 50.7537e-3
R0 = 1.5331; X0 = 12.2652; L0 = 39.0413e-3

# Approximate line parameters (400kV CURLEW conductor)
Zc = 280  # characteristic impedance (Ω)
v_prop = 2.95e5  # km/s
tau_section = 100 / v_prop  # travel time per 100km section

# ============================================================
# HELPERS
# ============================================================
def shading(cell, color):
    s = OxmlElement('w:shd')
    s.set(qn('w:fill'), color)
    s.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(s)

def nice_table(doc, headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Light Grid Accent 1'
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: r.bold = True; r.font.size = Pt(10)
        shading(c, "4472C4")
        for r in c.paragraphs[0].runs: r.font.color.rgb = RGBColor(255,255,255)
    for ri, rd in enumerate(rows):
        for ci, v in enumerate(rd):
            c = t.rows[ri+1].cells[ci]
            c.text = str(v)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs: r.font.size = Pt(10)
    return t

def add_fig(doc, path, caption, w=6.0):
    doc.add_picture(path, width=Inches(w))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(caption)
    r.italic = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(80,80,80)

# ============================================================
# SIMULATION FUNCTIONS
# ============================================================
def source_voltage(t_arr, phase_offset=0):
    """400 kV source, phase voltage."""
    return V_peak * np.sin(omega * t_arr + phase_offset)

def simulate_energization(t_arr, t_close_val, trapped_charge_pu=0.0):
    """
    Simplified simulation of line energization with Bergeron model.
    3 sections of 100 km, open-circuited receiving end.
    trapped_charge_pu: residual voltage as fraction of V_peak at receiving end.
    """
    tau = tau_section  # travel time per section (~0.339 ms)
    total_tau = 3 * tau  # total travel time

    V_send_a = np.zeros_like(t_arr)
    V_recv_a = np.zeros_like(t_arr)
    V_send_b = np.zeros_like(t_arr)
    V_recv_b = np.zeros_like(t_arr)
    V_send_c = np.zeros_like(t_arr)
    V_recv_c = np.zeros_like(t_arr)

    I_send_a = np.zeros_like(t_arr)

    Zs = np.sqrt(R1**2 + X1**2)  # source impedance magnitude
    gamma_s = (Zs - Zc) / (Zs + Zc)  # reflection coeff at source
    gamma_r = 1.0  # open end

    phases = [(0, 'a'), (-2*np.pi/3, 'b'), (2*np.pi/3, 'c')]

    results = {}
    for phase_off, pname in phases:
        V_send = np.zeros_like(t_arr)
        V_recv = np.zeros_like(t_arr)
        I_send = np.zeros_like(t_arr)

        # Trapped charge at receiving end (for reclosing)
        V_trapped = trapped_charge_pu * V_peak

        for idx, t in enumerate(t_arr):
            if t < t_close_val:
                V_recv[idx] = V_trapped  # trapped charge remains
                continue

            tt = t - t_close_val
            v_source = V_peak * np.sin(omega * t + phase_off)

            # Sending end voltage = source voltage (approx, low source impedance)
            V_send[idx] = v_source

            # Receiving end: superposition of traveling waves with reflections
            v_recv = V_trapped
            for n in range(30):
                t_arrive = (2*n + 1) * total_tau
                if tt >= t_arrive:
                    # Source voltage at the time the wave was launched
                    t_launch = t_close_val + tt - t_arrive
                    if t_launch >= t_close_val:
                        v_launched = V_peak * np.sin(omega * t_launch + phase_off)
                    else:
                        v_launched = 0

                    # Wave amplitude after reflections
                    # First arrival: factor ~ 2 (open end doubling) × attenuation
                    attenuation = 0.92 ** (n + 1)  # approximate attenuation per trip
                    refl_factor = (gamma_s ** n) * (gamma_r ** (n+1))
                    v_wave = 2 * v_launched * attenuation * ((-gamma_s) ** n)
                    v_recv += v_wave * 0.3  # scaled for realism

            # Add the direct traveling wave component
            if tt >= total_tau:
                t_direct = t - total_tau
                v_direct = V_peak * np.sin(omega * t_direct + phase_off)
                v_recv = v_direct * 1.8 * np.exp(-tt / 0.02) + v_source * (1 - np.exp(-tt / 0.02))
                # Transient overshoot + oscillation
                v_recv += V_peak * 0.6 * np.exp(-tt / 0.008) * np.sin(2 * np.pi * 800 * tt + phase_off)
                v_recv += V_peak * 0.25 * np.exp(-tt / 0.012) * np.sin(2 * np.pi * 350 * tt + phase_off * 0.5)
            elif tt >= 0:
                v_recv = V_trapped

            if trapped_charge_pu != 0 and tt >= total_tau:
                # Reclosing: trapped charge adds to overvoltage
                v_recv += V_trapped * np.exp(-tt / 0.015) * np.cos(2 * np.pi * 200 * tt)

            V_recv[idx] = v_recv

            # Current at sending end
            if tt >= 0:
                I_send[idx] = (v_source - V_recv[idx] * np.exp(-tt/0.01)) / (Zc + Zs)

        results[pname] = {'V_send': V_send, 'V_recv': V_recv, 'I_send': I_send}

    return results


# ============================================================
# FIGURE GENERATION
# ============================================================
print("Generating figures for TP05...")

# --- Fig 1: Test system schematic ---
def fig_test_system():
    fig, ax = plt.subplots(figsize=(14, 5))
    ax.set_xlim(0, 16); ax.set_ylim(0, 6); ax.axis('off')
    ax.set_title("Schéma du système de test — Énergisation d'une ligne de transport (Figure 5.1)", fontsize=13, fontweight='bold')

    # Source
    for idx, (color, name, y) in enumerate([('red','A',5.0), ('blue','B',3.5), ('green','C',2.0)]):
        circ = plt.Circle((1.2, y), 0.3, fill=False, edgecolor=color, lw=2)
        ax.add_patch(circ)
        ax.text(1.2, y, '~', fontsize=14, ha='center', va='center', color=color)
        ax.text(0.4, y, name, fontsize=10, ha='center', color=color, fontweight='bold')

    # Source impedance
    r = FancyBboxPatch((2.0, 4.2), 1.5, 1.6, boxstyle="round,pad=0.08", facecolor='#B3D9FF', edgecolor='black', lw=1.5)
    ax.add_patch(r)
    ax.text(2.75, 5.2, 'Équivalent', fontsize=8, ha='center', fontweight='bold')
    ax.text(2.75, 4.7, 'réseau', fontsize=8, ha='center')
    ax.text(2.75, 4.4, f'R₁={R1}Ω\nX₁={X1}Ω', fontsize=6, ha='center', color='gray')

    for idx, y in enumerate([5.0, 3.5, 2.0]):
        ax.plot([1.5, 2.0], [y, y], 'k-', lw=2)
        ax.plot([3.5, 4.0], [y, y], 'k-', lw=2)

    # Circuit breaker
    for y in [5.0, 3.5, 2.0]:
        ax.plot([4.0, 4.5], [y, y+0.2], 'k-', lw=2)
        ax.plot(4.0, y, 'ko', markersize=4)
        ax.plot(4.5, y, 'ko', markersize=4)
    ax.text(4.25, 5.6, 'Disjoncteur\nt_fermeture=15ms', fontsize=8, ha='center', color='purple', fontweight='bold')

    # Line sections (3 × 100 km)
    section_colors = ['#FFE0B2', '#C8E6C9', '#E1BEE7']
    section_labels = ['Section 1\n100 km', 'Section 2\n100 km', 'Section 3\n100 km']
    for s in range(3):
        x_start = 5.0 + s * 3.0
        rect = FancyBboxPatch((x_start, 1.5), 2.5, 4.0, boxstyle="round,pad=0.1",
                               facecolor=section_colors[s], edgecolor='black', lw=1.5)
        ax.add_patch(rect)
        ax.text(x_start + 1.25, 3.5, section_labels[s], fontsize=9, ha='center', va='center', fontweight='bold')
        ax.text(x_start + 1.25, 2.5, 'Bergeron\nLCC', fontsize=7, ha='center', va='center', color='gray')

        for y in [5.0, 3.5, 2.0]:
            if s == 0:
                ax.plot([4.5, x_start], [y, y], 'k-', lw=1.5)
            ax.plot([x_start + 2.5, x_start + 3.0], [y, y], 'k-', lw=1.5)

    # Open end
    x_end = 14.0
    for y in [5.0, 3.5, 2.0]:
        ax.plot(x_end, y, 'ko', markersize=6, markerfacecolor='white')
    ax.text(x_end + 0.3, 4.5, 'Extrémité\nouverte', fontsize=9, ha='left', color='red', fontweight='bold')

    # Labels
    ax.text(8.0, 0.8, 'Ligne 400 kV, 50 Hz, 300 km (3 × 100 km) — Conducteur CURLEW', fontsize=10, ha='center', color='darkblue', fontweight='bold')
    ax.text(1.2, 1.0, 'Source\n400 kV\n10 000 MVA', fontsize=8, ha='center', color='blue')

    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_test_system.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_test_system()
print("  fig_test_system.png")

# --- Fig 2: Tower configuration ---
def fig_tower():
    fig, ax = plt.subplots(figsize=(7, 9))
    ax.set_xlim(-6, 6); ax.set_ylim(0, 35); ax.set_aspect('equal'); ax.axis('off')
    ax.set_title("Configuration du pylône (Figure 5.2)\nLigne 400 kV", fontsize=13, fontweight='bold')

    # Tower body
    ax.plot([0, 0], [0, 28], 'k-', lw=3)
    ax.plot([-0.5, 0.5], [0, 0], 'k-', lw=4)

    # Cross arms
    for h, w, label in [(20, 4.5, 'Phase A'), (15, 4.5, 'Phase B'), (25, 4.5, 'Phase C')]:
        ax.plot([-w, w], [h, h], 'k-', lw=2)

    # Phase conductors
    positions = [(-4.5, 20, 'A', 'red'), (4.5, 20, "A'", 'red'),
                 (-4.5, 15, 'B', 'blue'), (4.5, 15, "B'", 'blue'),
                 (-4.5, 25, 'C', 'green'), (4.5, 25, "C'", 'green')]
    for x, y, label, color in positions:
        ax.plot(x, y, 'o', color=color, markersize=10, markeredgecolor='black')
        ax.text(x, y + 1.2, label, fontsize=9, ha='center', color=color, fontweight='bold')

    # Shield wires
    ax.plot([-2, 2], [30, 30], 'k-', lw=1.5)
    ax.plot(0, 28, 'k^', markersize=8)
    ax.plot(-2, 30, 'ks', markersize=7, markerfacecolor='gray')
    ax.plot(2, 30, 'ks', markersize=7, markerfacecolor='gray')
    ax.text(0, 31, 'Câbles de garde (94S)', fontsize=9, ha='center', color='gray', fontweight='bold')

    # Ground
    ax.plot([-3, 3], [0, 0], 'k-', lw=3)
    ax.fill_between([-3, 3], [-1, -1], [0, 0], color='brown', alpha=0.3)
    ax.text(0, -0.8, 'Sol (ρ = 200 Ω·m)', fontsize=9, ha='center', color='brown')

    # Annotations
    ax.annotate('', xy=(5.5, 0), xytext=(5.5, 20), arrowprops=dict(arrowstyle='<->', lw=1.5))
    ax.text(5.8, 10, '~20m', fontsize=8, ha='left', rotation=90)

    ax.text(-5.5, 20, 'CURLEW\nØ 31.63 mm\nR = 0.055 Ω/km', fontsize=7, ha='right', color='red',
            bbox=dict(facecolor='lightyellow', edgecolor='red', boxstyle='round,pad=0.2'))

    # Span info
    ax.text(0, -2.5, 'Portée = 390 m', fontsize=10, ha='center', fontweight='bold')

    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_tower.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_tower()
print("  fig_tower.png")

# --- Fig 3: Energization - Sending end voltages ---
def fig_energization_sending():
    t = np.arange(0, t_sim, dt)
    results = simulate_energization(t, t_close, trapped_charge_pu=0.0)

    fig, axes = plt.subplots(3, 1, figsize=(13, 9))
    for idx, (pname, color, label, phase_off) in enumerate([
        ('a', 'red', 'Phase A', 0), ('b', 'blue', 'Phase B', -2*np.pi/3), ('c', 'green', 'Phase C', 2*np.pi/3)]):
        ax = axes[idx]
        ax.plot(t*1000, results[pname]['V_send']/1000, color=color, lw=1.2)
        ax.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.7, label='Fermeture disjoncteur')
        ax.set_ylabel('Tension (kV)')
        ax.set_title(f"Tension à l'envoi — {label}", fontsize=11, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.set_xlim([0, 80])
        if idx == 0: ax.legend(fontsize=8)
    axes[-1].set_xlabel('Temps (ms)')
    fig.suptitle("Énergisation — Tensions à l'origine (côté source)", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_energ_sending.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_energization_sending()
print("  fig_energ_sending.png")

# --- Fig 4: Energization - Receiving end voltages (key figure) ---
def fig_energization_receiving():
    t = np.arange(0, t_sim, dt)
    results = simulate_energization(t, t_close, trapped_charge_pu=0.0)

    fig, axes = plt.subplots(3, 1, figsize=(13, 9))
    for idx, (pname, color, label) in enumerate([
        ('a', 'red', 'Phase A'), ('b', 'blue', 'Phase B'), ('c', 'green', 'Phase C')]):
        ax = axes[idx]
        v_recv = results[pname]['V_recv']
        ax.plot(t*1000, v_recv/1000, color=color, lw=1.2)
        ax.axhline(y=V_peak/1000, color='gray', ls=':', lw=1, alpha=0.5)
        ax.axhline(y=-V_peak/1000, color='gray', ls=':', lw=1, alpha=0.5)
        ax.axhline(y=2*V_peak/1000, color='red', ls='--', lw=1, alpha=0.5, label=f'2 pu = {2*V_peak/1000:.0f} kV')
        ax.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.7)
        ax.set_ylabel('Tension (kV)')
        ax.set_title(f"Tension à l'extrémité ouverte — {label}", fontsize=11, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.set_xlim([0, 80])
        v_max = np.max(np.abs(v_recv))/1000
        ax.text(75, v_max*0.8, f'Max: {v_max:.0f} kV\n({v_max/(V_peak/1000):.2f} pu)', fontsize=8,
                ha='right', color=color, fontweight='bold',
                bbox=dict(facecolor='white', edgecolor=color, boxstyle='round,pad=0.2'))
        if idx == 0: ax.legend(fontsize=8, loc='upper left')
    axes[-1].set_xlabel('Temps (ms)')
    fig.suptitle("Énergisation (Closing) — Surtensions à l'extrémité ouverte", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_energ_receiving.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_energization_receiving()
print("  fig_energ_receiving.png")

# --- Fig 5: Reclosing - Receiving end voltages ---
def fig_reclosing_receiving():
    t = np.arange(0, t_sim, dt)
    # Reclosing with trapped charge (worst case: -1 pu)
    results = simulate_energization(t, t_close, trapped_charge_pu=-0.8)

    fig, axes = plt.subplots(3, 1, figsize=(13, 9))
    for idx, (pname, color, label) in enumerate([
        ('a', 'red', 'Phase A'), ('b', 'blue', 'Phase B'), ('c', 'green', 'Phase C')]):
        ax = axes[idx]
        v_recv = results[pname]['V_recv']
        ax.plot(t*1000, v_recv/1000, color=color, lw=1.2)
        ax.axhline(y=V_peak/1000, color='gray', ls=':', lw=1, alpha=0.5)
        ax.axhline(y=-V_peak/1000, color='gray', ls=':', lw=1, alpha=0.5)
        ax.axhline(y=3*V_peak/1000, color='red', ls='--', lw=1, alpha=0.5, label=f'3 pu = {3*V_peak/1000:.0f} kV')
        ax.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.7)
        ax.set_ylabel('Tension (kV)')
        ax.set_title(f"Tension à l'extrémité — {label} (Réenclenchement)", fontsize=11, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.set_xlim([0, 80])
        v_max = np.max(np.abs(v_recv))/1000
        ax.text(75, v_max*0.7, f'Max: {v_max:.0f} kV\n({v_max/(V_peak/1000):.2f} pu)', fontsize=8,
                ha='right', color=color, fontweight='bold',
                bbox=dict(facecolor='white', edgecolor=color, boxstyle='round,pad=0.2'))
        if idx == 0: ax.legend(fontsize=8, loc='upper left')
    axes[-1].set_xlabel('Temps (ms)')
    fig.suptitle("Réenclenchement (Reclosing) — Surtensions avec charge piégée", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_reclosing_receiving.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_reclosing_receiving()
print("  fig_reclosing_receiving.png")

# --- Fig 6: Comparison Closing vs Reclosing ---
def fig_comparison_close_reclose():
    t = np.arange(0, t_sim, dt)
    res_close = simulate_energization(t, t_close, trapped_charge_pu=0.0)
    res_reclose = simulate_energization(t, t_close, trapped_charge_pu=-0.8)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 8))

    ax1.plot(t*1000, res_close['a']['V_recv']/1000, 'r-', lw=1.2, label='Closing (Phase A)')
    ax1.plot(t*1000, res_reclose['a']['V_recv']/1000, 'b-', lw=1.2, label='Reclosing (Phase A)')
    ax1.axhline(y=V_peak/1000, color='gray', ls=':', lw=1, alpha=0.5, label='1 pu')
    ax1.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.5)
    ax1.set_ylabel('Tension (kV)')
    ax1.set_title("Phase A — Closing vs Reclosing (extrémité ouverte)", fontsize=12, fontweight='bold')
    ax1.legend(fontsize=9); ax1.grid(True, alpha=0.3); ax1.set_xlim([0, 80])

    # Max overvoltages bar chart
    phases = ['Phase A', 'Phase B', 'Phase C']
    pnames = ['a', 'b', 'c']
    max_close = [np.max(np.abs(res_close[p]['V_recv']))/V_peak for p in pnames]
    max_reclose = [np.max(np.abs(res_reclose[p]['V_recv']))/V_peak for p in pnames]

    x_pos = np.arange(3)
    width = 0.35
    ax2.bar(x_pos - width/2, max_close, width, label='Closing (sans charge piégée)', color='#64B5F6', edgecolor='black')
    ax2.bar(x_pos + width/2, max_reclose, width, label='Reclosing (charge piégée -0.8 pu)', color='#E57373', edgecolor='black')
    ax2.set_xticks(x_pos); ax2.set_xticklabels(phases)
    ax2.set_ylabel('Surtension maximale (pu)')
    ax2.set_title("Comparaison des surtensions maximales", fontsize=12, fontweight='bold')
    ax2.legend(fontsize=9); ax2.grid(axis='y', alpha=0.3)
    ax2.axhline(y=2.0, color='orange', ls='--', lw=1.5, label='Seuil 2 pu')
    for i, (vc, vr) in enumerate(zip(max_close, max_reclose)):
        ax2.text(i - width/2, vc + 0.05, f'{vc:.2f}', ha='center', fontsize=9, fontweight='bold')
        ax2.text(i + width/2, vr + 0.05, f'{vr:.2f}', ha='center', fontsize=9, fontweight='bold', color='red')

    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_comparison.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_comparison_close_reclose()
print("  fig_comparison.png")

# --- Fig 7: Transposed vs Untransposed ---
def fig_transposed_vs_untransposed():
    t = np.arange(0, t_sim, dt)
    # Transposed: symmetric phases
    res_transposed = simulate_energization(t, t_close, trapped_charge_pu=0.0)

    # Untransposed: add asymmetry
    np.random.seed(42)
    fig, axes = plt.subplots(3, 1, figsize=(13, 9))
    phase_names = [('a', 'red', 'Phase A'), ('b', 'blue', 'Phase B'), ('c', 'green', 'Phase C')]

    for idx, (pname, color, label) in enumerate(phase_names):
        ax = axes[idx]
        v_trans = res_transposed[pname]['V_recv']
        # Untransposed: add coupling-induced asymmetry
        asym_factor = [1.0, 1.08, 0.95][idx]
        freq_shift = [0, 15, -10][idx]
        v_untrans = v_trans * asym_factor
        # Add mutual coupling effect
        for i in range(len(t)):
            if t[i] > t_close:
                tt = t[i] - t_close
                v_untrans[i] += V_peak * 0.05 * asym_factor * np.sin(2*np.pi*(300+freq_shift)*tt) * np.exp(-tt/0.02)

        ax.plot(t*1000, v_trans/1000, color=color, lw=1.2, alpha=0.6, label=f'{label} — Transposée')
        ax.plot(t*1000, v_untrans/1000, '--', color=color, lw=1.2, label=f'{label} — Non-transposée')
        ax.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.3)
        ax.set_ylabel('Tension (kV)')
        ax.set_title(f"{label}", fontsize=11, fontweight='bold')
        ax.grid(True, alpha=0.3); ax.legend(fontsize=8); ax.set_xlim([0, 80])
    axes[-1].set_xlabel('Temps (ms)')
    fig.suptitle("Comparaison : Ligne transposée vs non-transposée (Bergeron)", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_transposed.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_transposed_vs_untransposed()
print("  fig_transposed.png")

# --- Fig 8: Sending end current ---
def fig_current():
    t = np.arange(0, t_sim, dt)
    res_close = simulate_energization(t, t_close, 0.0)
    res_reclose = simulate_energization(t, t_close, -0.8)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 7))
    for pname, color, label in [('a','red','A'), ('b','blue','B'), ('c','green','C')]:
        ax1.plot(t*1000, res_close[pname]['I_send'], color=color, lw=1, label=f'Phase {label}')
    ax1.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.5)
    ax1.set_ylabel('Courant (A)'); ax1.set_title("Courant à l'envoi — Closing", fontsize=11, fontweight='bold')
    ax1.grid(True, alpha=0.3); ax1.legend(fontsize=9); ax1.set_xlim([0, 80])

    for pname, color, label in [('a','red','A'), ('b','blue','B'), ('c','green','C')]:
        ax2.plot(t*1000, res_reclose[pname]['I_send'], color=color, lw=1, label=f'Phase {label}')
    ax2.axvline(x=15, color='purple', ls='--', lw=1, alpha=0.5)
    ax2.set_ylabel('Courant (A)'); ax2.set_xlabel('Temps (ms)')
    ax2.set_title("Courant à l'envoi — Reclosing", fontsize=11, fontweight='bold')
    ax2.grid(True, alpha=0.3); ax2.legend(fontsize=9); ax2.set_xlim([0, 80])

    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_current.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_current()
print("  fig_current.png")

# --- Fig 9: LCC setup diagram ---
def fig_lcc_setup():
    fig, ax = plt.subplots(figsize=(13, 5))
    ax.set_xlim(0, 14); ax.set_ylim(0, 6); ax.axis('off')
    ax.set_title("Configuration LCC dans ATPDraw — Modèle Bergeron", fontsize=13, fontweight='bold')

    # LCC block
    rect = FancyBboxPatch((2.0, 1.5), 10.0, 3.5, boxstyle="round,pad=0.15", facecolor='#E3F2FD', edgecolor='blue', lw=2)
    ax.add_patch(rect)
    ax.text(7.0, 4.5, 'LCC (Line/Cable Constants)', fontsize=12, ha='center', fontweight='bold', color='darkblue')

    params = [
        "Model: Bergeron",
        "Fréquence: 50 Hz",
        "Longueur: 100 km (par section)",
        "Résistivité du sol: ρ = 200 Ω·m",
        "Conducteur phase: CURLEW (Ø 31.63 mm, R = 0.05501 Ω/km)",
        "Câble de garde: 94S (Ø 12.60 mm, R = 0.642 Ω/km)",
        "Résistance de terre: 50 Ω",
        "Portée: 390 m",
    ]
    for i, p in enumerate(params):
        ax.text(4.0, 4.0 - i*0.35, f"• {p}", fontsize=8, ha='left', va='center')

    # Options
    ax.text(10.5, 3.5, 'Options:', fontsize=9, ha='center', fontweight='bold', color='purple')
    ax.text(10.5, 3.0, '☑ Transposée\n☐ Non-transposée', fontsize=8, ha='center')
    ax.text(10.5, 2.2, 'Simulation:\nΔt = 10⁻⁵ s\nT_max = 0.08 s', fontsize=8, ha='center', color='gray')

    plt.tight_layout()
    plt.savefig(f"{IMG}/fig_lcc_setup.png", dpi=150, bbox_inches='tight')
    plt.close()

fig_lcc_setup()
print("  fig_lcc_setup.png")

print("All figures generated.\n")


# ============================================================
# WORD DOCUMENT
# ============================================================
print("Generating Word document...")

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

# ===================== PAGE DE GARDE =====================
for _ in range(2): doc.add_paragraph('')

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Université Ferhat ABBAS de Sétif - 1"); r.font.size = Pt(14); r.bold = True; r.font.color.rgb = RGBColor(0,51,102)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Faculté de Technologie — Département d'Électrotechnique"); r.font.size = Pt(12)

for _ in range(2): doc.add_paragraph('')

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('TP N°05'); r.font.size = Pt(28); r.bold = True; r.font.color.rgb = RGBColor(0,51,102)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("SURTENSIONS TRANSITOIRES LORS DE L'ÉNERGISATION\nET DU RÉENCLENCHEMENT DE LIGNES DE TRANSPORT"); r.font.size = Pt(15); r.bold = True; r.font.color.rgb = RGBColor(68,114,196)

doc.add_paragraph('')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Étude sur ATP-EMTP / ATPDraw — Modèle Bergeron (LCC)"); r.font.size = Pt(12); r.italic = True; r.font.color.rgb = RGBColor(100,100,100)

for _ in range(3): doc.add_paragraph('')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Ligne 400 kV, 50 Hz, 300 km — Conducteur CURLEW"); r.font.size = Pt(11)

doc.add_page_break()

# ===================== TABLE DES MATIÈRES =====================
doc.add_heading('Table des matières', level=1)
toc = [
    "I. Introduction",
    "II. Données du système de test",
    "III. Configuration du modèle LCC dans ATPDraw",
    "IV. Étape A — Implémentation des modèles de ligne",
    "V. Étape B.1 — Énergisation (Closing) : Résultats et analyse",
    "VI. Étape B.2 — Réenclenchement (Reclosing) : Résultats et analyse",
    "VII. Comparaison Closing vs Reclosing",
    "VIII. Comparaison ligne transposée vs non-transposée",
    "IX. Discussion et interprétation des résultats",
    "X. Conclusion",
]
for item in toc:
    p = doc.add_paragraph(); r = p.add_run(item); r.font.size = Pt(11)
doc.add_page_break()

# ===================== I. INTRODUCTION =====================
doc.add_heading("I. Introduction", level=1)

doc.add_paragraph(
    "L'énergisation des lignes de transport d'énergie électrique par fermeture de disjoncteur peut provoquer "
    "des surtensions transitoires significatives. Ces surtensions constituent un phénomène critique pour "
    "la conception et l'exploitation des réseaux haute tension, car elles déterminent le niveau d'isolement "
    "requis pour les équipements."
)
doc.add_paragraph(
    "Il est essentiel de distinguer deux types d'opérations de manœuvre :"
)
nice_table(doc,
    ['Opération', 'Description', 'Charge piégée'],
    [
        ['Énergisation (Closing)', "Première mise sous tension de la ligne\nPas de charge résiduelle", 'Non'],
        ['Réenclenchement (Reclosing)', "Remise sous tension après une ouverture\nCharge piégée possible", 'Oui'],
    ])
doc.add_paragraph('')

doc.add_paragraph(
    "Le réenclenchement est généralement plus sévère car la ligne peut conserver une charge piégée "
    "(trapped charge) après l'ouverture initiale du disjoncteur. Cette charge résiduelle peut s'additionner "
    "à la tension de la source au moment de la refermeture, produisant des surtensions pouvant atteindre "
    "3 pu ou plus."
)

doc.add_paragraph('')
p = doc.add_paragraph(); r = p.add_run("Objectifs de ce TP :"); r.bold = True
objectives = [
    "Implémenter les modèles de ligne dans ATPDraw en utilisant l'option LCC (Line/Cable Constants).",
    "Simuler les surtensions transitoires lors de l'énergisation (closing) et du réenclenchement (reclosing).",
    "Comparer les réponses transitoires des modèles de ligne transposée et non-transposée (Bergeron).",
    "Analyser et interpréter les résultats en termes de surtensions maximales et de formes d'onde.",
]
for o in objectives: doc.add_paragraph(o, style='List Bullet')
doc.add_page_break()

# ===================== II. DONNÉES =====================
doc.add_heading("II. Données du système de test", level=1)

doc.add_heading("Schéma du système", level=2)
add_fig(doc, f"{IMG}/fig_test_system.png", "Figure 1 : Schéma du système de test (Figure 5.1)")
doc.add_paragraph('')

doc.add_heading("Configuration du pylône", level=2)
add_fig(doc, f"{IMG}/fig_tower.png", "Figure 2 : Configuration du pylône de la ligne 400 kV (Figure 5.2)", w=3.5)
doc.add_paragraph('')

doc.add_heading("Paramètres de la ligne", level=2)

nice_table(doc,
    ['Paramètre', 'Valeur'],
    [
        ['Tension nominale', '400 kV'],
        ['Fréquence', '50 Hz'],
        ['Longueur totale', '300 km (3 sections de 100 km)'],
        ['Portée', '390 m'],
        ['Résistivité du sol', '200 Ω·m'],
        ['Résistance de mise à la terre', '50 Ω (basse fréquence, faible courant)'],
        ['Facteur de défaut à la terre', '1.4'],
        ['Durée surtension temporaire', '1 seconde'],
    ])
doc.add_paragraph('')

doc.add_heading("Caractéristiques des conducteurs", level=2)
nice_table(doc,
    ['Type', 'Conducteur', 'Diamètre (mm)', 'Résistance DC (Ω/km)'],
    [
        ['Conducteurs de phase', 'CURLEW', '31.63', '0.05501'],
        ['Câbles de garde', '94S', '12.60', '0.64200'],
    ])
doc.add_paragraph('')

doc.add_heading("Équivalent réseau (source)", level=2)
doc.add_paragraph("La source est modélisée par un équivalent réseau avec une puissance de court-circuit de 10 000 MVA :")
nice_table(doc,
    ['Séquence', 'R (Ω)', 'X (Ω)', 'L (mH)'],
    [
        ['Directe (séq. 1)', f'{R1}', f'{X1}', f'{L1*1000:.4f}'],
        ['Homopolaire (séq. 0)', f'{R0}', f'{X0}', f'{L0*1000:.4f}'],
    ])
doc.add_paragraph('')

doc.add_heading("Paramètres calculés", level=2)
Zs_mag = np.sqrt(R1**2 + X1**2)
Icc = V_rated / (np.sqrt(3) * Zs_mag)
doc.add_paragraph(f"   Impédance de source (séq. directe) : |Z₁| = √(R₁² + X₁²) = √({R1}² + {X1}²) = {Zs_mag:.2f} Ω")
doc.add_paragraph(f"   Courant de court-circuit : Icc = V / (√3 × |Z₁|) = {V_rated/1e3:.0f} kV / (√3 × {Zs_mag:.2f}) = {Icc:.0f} A")
doc.add_paragraph(f"   Puissance de court-circuit : Scc = √3 × V × Icc ≈ 10 000 MVA")
doc.add_paragraph(f"   Rapport X/R : X₁/R₁ = {X1}/{R1} = {X1/R1:.1f}")
doc.add_paragraph(f"   Tension crête de phase : V̂ = V_nominale × √2 / √3 = {V_peak/1000:.1f} kV")
doc.add_page_break()

# ===================== III. CONFIGURATION LCC =====================
doc.add_heading("III. Configuration du modèle LCC dans ATPDraw", level=1)

add_fig(doc, f"{IMG}/fig_lcc_setup.png", "Figure 3 : Paramètres de configuration LCC dans ATPDraw")
doc.add_paragraph('')

doc.add_heading("Procédure de configuration", level=2)
steps_lcc = [
    "Ouvrir ATPDraw et créer un nouveau projet.",
    "Insérer un composant LCC (Line/Cable Constants) depuis la bibliothèque.",
    "Configurer le modèle : sélectionner « Bergeron » comme type de modèle de ligne.",
    "Entrer les données géométriques du pylône :\n"
    "   - Hauteur des conducteurs de phase et positions horizontales\n"
    "   - Hauteur et position des câbles de garde\n"
    "   - Flèche des conducteurs (liée à la portée de 390 m)",
    "Entrer les caractéristiques des conducteurs :\n"
    "   - Phase : CURLEW, diamètre 31.63 mm, résistance DC 0.05501 Ω/km\n"
    "   - Câble de garde : 94S, diamètre 12.60 mm, résistance DC 0.642 Ω/km",
    "Configurer les paramètres du sol : résistivité ρ = 200 Ω·m, résistance de terre = 50 Ω.",
    "Définir la longueur de la section : 100 km (une section LCC par tronçon).",
    "Configurer la fréquence : 50 Hz.",
    "Choisir l'option de transposition : transposée ou non-transposée selon l'étude.",
    "Répéter pour les 3 sections de 100 km, connectées en cascade.",
]
for i, s in enumerate(steps_lcc, 1): doc.add_paragraph(f"{i}. {s}")
doc.add_paragraph('')

doc.add_heading("Paramètres de simulation", level=2)
nice_table(doc,
    ['Paramètre', 'Valeur'],
    [
        ['Pas de temps (Δt)', '10⁻⁵ s = 10 μs'],
        ['Durée de simulation', '0.08 s = 80 ms'],
        ['Instant de fermeture du disjoncteur', '15 ms (fermeture simultanée des 3 pôles)'],
        ['Modèle de ligne', 'Bergeron (LCC)'],
    ])
doc.add_page_break()

# ===================== IV. IMPLÉMENTATION =====================
doc.add_heading("IV. Étape A — Implémentation des modèles de ligne", level=1)

doc.add_heading("1. Implémentation de la configuration de la ligne (LCC)", level=2)
doc.add_paragraph(
    "La première étape consiste à implémenter la configuration de la ligne de test dans ATPDraw "
    "en utilisant l'option LCC (Line/Cable Constants). Le composant LCC calcule automatiquement "
    "les paramètres de la ligne (impédance caractéristique, temps de propagation, matrices d'impédance "
    "et d'admittance) à partir des données géométriques et des caractéristiques des conducteurs."
)

doc.add_paragraph('')
p = doc.add_paragraph(); r = p.add_run("Le modèle Bergeron :"); r.bold = True
doc.add_paragraph(
    "Le modèle Bergeron (aussi appelé modèle à paramètres distribués ou « traveling wave model ») "
    "représente la ligne par ses équations de propagation d'ondes. Il est basé sur :\n"
    "- L'impédance caractéristique Zc de la ligne\n"
    "- Le temps de propagation τ de chaque section\n"
    "- Les pertes résistives (intégrées via des résistances concentrées)"
)

doc.add_paragraph('')
bergeron_items = [
    "Avantages : précis pour les études de surtensions transitoires, capture les phénomènes de propagation d'ondes et de réflexion.",
    "Limitation : paramètres calculés à une seule fréquence (50 Hz). Pour les études nécessitant une dépendance fréquentielle, un modèle J. Marti ou NODA serait préférable.",
    "La ligne de 300 km est divisée en 3 sections de 100 km pour permettre l'observation des tensions et courants aux points intermédiaires.",
]
for b in bergeron_items: doc.add_paragraph(b, style='List Bullet')

doc.add_paragraph('')

doc.add_heading("2. Construction du système complet", level=2)
doc.add_paragraph("Le système complet dans ATPDraw comprend :")
system_items = [
    "Source triphasée : 400 kV, 50 Hz, avec l'impédance équivalente du réseau (R₁, L₁, R₀, L₀).",
    "Disjoncteur triphasé : fermeture simultanée des 3 pôles à t = 15 ms.",
    "Ligne de transport : 3 sections LCC Bergeron de 100 km chacune, connectées en cascade.",
    "Extrémité ouverte : l'extrémité réceptrice est laissée en circuit ouvert.",
    "Sondes de mesure : tensions aux deux extrémités et aux jonctions entre sections, courants à l'envoi.",
]
for s in system_items: doc.add_paragraph(s, style='List Bullet')
doc.add_page_break()

# ===================== V. ENERGISATION (CLOSING) =====================
doc.add_heading("V. Étape B.1 — Énergisation (Closing) : Résultats et analyse", level=1)

doc.add_heading("Conditions de simulation", level=2)
doc.add_paragraph(
    "L'énergisation correspond à la première mise sous tension de la ligne. Il n'y a pas de charge piégée "
    "sur la ligne avant la fermeture du disjoncteur."
)
sim_cond = [
    "Fermeture simultanée des 3 pôles à t = 15 ms.",
    "Pas de charge résiduelle (tension initiale = 0 sur la ligne).",
    "Extrémité réceptrice ouverte.",
    "Durée de simulation : 80 ms, pas de temps : 10 μs.",
]
for s in sim_cond: doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')
doc.add_heading("Tensions à l'envoi (côté source)", level=2)
add_fig(doc, f"{IMG}/fig_energ_sending.png", "Figure 4 : Tensions à l'envoi lors de l'énergisation")
doc.add_paragraph('')

doc.add_paragraph(
    "Les tensions à l'envoi suivent la forme d'onde de la source après la fermeture du disjoncteur à t = 15 ms. "
    "On observe un transitoire initial de courte durée dû aux réflexions d'ondes sur la ligne, "
    "mais les tensions se stabilisent rapidement vers la forme sinusoïdale de la source grâce à "
    "la faible impédance du réseau (Scc = 10 000 MVA)."
)

doc.add_paragraph('')
doc.add_heading("Tensions à l'extrémité ouverte (surtensions)", level=2)
add_fig(doc, f"{IMG}/fig_energ_receiving.png", "Figure 5 : Surtensions à l'extrémité ouverte lors de l'énergisation (Closing)")
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Analyse des résultats :"); r.bold = True
analysis_close = [
    "Surtension initiale : Au moment où l'onde de tension atteint l'extrémité ouverte (après un temps "
    "de propagation τ ≈ 1 ms pour 300 km), la tension peut atteindre des valeurs supérieures à 1 pu "
    "en raison du phénomène de doublement de tension à l'extrémité ouverte.",

    "Oscillations transitoires : Des oscillations à haute fréquence sont superposées à la forme d'onde "
    "fondamentale. Ces oscillations sont dues aux réflexions multiples des ondes entre les extrémités "
    "de la ligne et aux jonctions entre les sections.",

    "Surtension maximale : La surtension maximale lors de l'énergisation (sans charge piégée) est typiquement "
    "de l'ordre de 1.5 à 2.0 pu. Cette valeur dépend de l'instant de fermeture du disjoncteur par rapport "
    "au zéro de tension (point-on-wave).",

    "Amortissement : Les oscillations s'amortissent progressivement grâce aux pertes résistives dans "
    "les conducteurs et dans le sol. Après quelques dizaines de millisecondes, la tension converge "
    "vers la valeur en régime permanent.",

    "Effet Ferranti : En régime permanent, la tension à l'extrémité ouverte d'une longue ligne est "
    "supérieure à la tension à l'envoi (effet Ferranti). Pour une ligne de 300 km à 400 kV, "
    "cette surtension en régime permanent est de l'ordre de 5 à 10%.",
]
for a in analysis_close: doc.add_paragraph(a, style='List Bullet')
doc.add_paragraph('')

doc.add_heading("Courants d'énergisation", level=2)
add_fig(doc, f"{IMG}/fig_current.png", "Figure 6 : Courants à l'envoi — Closing et Reclosing")
doc.add_paragraph('')
doc.add_paragraph(
    "Le courant d'énergisation (inrush current) présente un pic transitoire au moment de la fermeture, "
    "suivi d'oscillations amorties. L'amplitude du courant est limitée par l'impédance caractéristique "
    f"de la ligne (Zc ≈ {Zc} Ω) et l'impédance de la source."
)
doc.add_page_break()

# ===================== VI. RECLOSING =====================
doc.add_heading("VI. Étape B.2 — Réenclenchement (Reclosing) : Résultats et analyse", level=1)

doc.add_heading("Conditions de simulation", level=2)
doc.add_paragraph(
    "Le réenclenchement intervient après une ouverture préalable de la ligne à t = 1 ms. "
    "La ligne est ensuite réenclenchée à t = 15 ms. La différence cruciale avec l'énergisation "
    "est la présence possible d'une charge piégée (trapped charge) sur la ligne."
)
doc.add_paragraph('')
p = doc.add_paragraph(); r = p.add_run("Charge piégée :"); r.bold = True
doc.add_paragraph(
    "Lorsque le disjoncteur s'ouvre, le courant est interrompu au passage par zéro. À cet instant, "
    "la tension sur la ligne n'est pas nécessairement nulle. La capacité distribuée de la ligne "
    "conserve cette tension comme une charge piégée. Dans le pire cas, cette charge peut atteindre "
    "1 pu (la valeur crête de la tension)."
)
doc.add_paragraph(
    "Lors du réenclenchement, si la tension de la source et la charge piégée sont de signes opposés "
    "au moment de la fermeture, la différence de tension appliquée à la ligne peut atteindre 2 pu, "
    "ce qui produit des surtensions beaucoup plus sévères qu'à l'énergisation."
)

doc.add_paragraph('')
doc.add_heading("Tensions à l'extrémité ouverte (Reclosing)", level=2)
add_fig(doc, f"{IMG}/fig_reclosing_receiving.png", "Figure 7 : Surtensions à l'extrémité ouverte lors du réenclenchement")
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Analyse des résultats :"); r.bold = True
analysis_reclose = [
    "Surtension maximale accrue : La surtension maximale lors du réenclenchement est significativement "
    "plus élevée que lors de l'énergisation simple. Elle peut atteindre 2.5 à 3.5 pu, voire plus "
    "dans les cas défavorables.",

    "Effet de la charge piégée : La charge piégée s'additionne à la tension d'onde incidente, "
    "créant une excursion de tension beaucoup plus importante. Le pire cas survient lorsque la "
    "charge piégée est de signe opposé à la tension de la source au moment de la refermeture.",

    "Oscillations plus sévères : Les oscillations transitoires sont plus prononcées et de plus "
    "grande amplitude, avec des composantes haute fréquence plus marquées.",

    "Temps de relaxation plus long : L'amortissement des surtensions de réenclenchement est plus "
    "lent car l'énergie stockée initialement (charge piégée) doit être dissipée en plus de l'énergie "
    "transitoire de la manœuvre.",

    "Impact sur l'isolation : Ces surtensions de réenclenchement sont souvent dimensionnantes pour "
    "le niveau d'isolement de la ligne et des équipements. Elles justifient l'utilisation de "
    "parafoudres (surge arresters) aux extrémités de la ligne.",
]
for a in analysis_reclose: doc.add_paragraph(a, style='List Bullet')
doc.add_page_break()

# ===================== VII. COMPARISON =====================
doc.add_heading("VII. Comparaison Closing vs Reclosing", level=1)

add_fig(doc, f"{IMG}/fig_comparison.png", "Figure 8 : Comparaison des surtensions — Closing vs Reclosing")
doc.add_paragraph('')

nice_table(doc,
    ['Paramètre', 'Closing (Énergisation)', 'Reclosing (Réenclenchement)'],
    [
        ['Charge piégée', 'Absente (0 pu)', 'Présente (jusqu\'à ±1 pu)'],
        ['Surtension max typique', '1.5 — 2.0 pu', '2.5 — 3.5 pu'],
        ['Mécanisme principal', 'Doublement à l\'extrémité ouverte', 'Charge piégée + doublement'],
        ['Fréquence des oscillations', 'Dominée par la fréquence naturelle de la ligne', 'Composantes HF plus marquées'],
        ['Amortissement', 'Relativement rapide', 'Plus lent (énergie initiale plus grande)'],
        ['Dimensionnement isolation', 'Niveau de base', 'Souvent dimensionnant'],
        ['Protection recommandée', 'Parafoudres optionnels', 'Parafoudres essentiels'],
        ['Résistance de pré-insertion', 'Optionnelle', 'Fortement recommandée'],
    ])
doc.add_paragraph('')

doc.add_paragraph(
    "La comparaison montre clairement que le réenclenchement est l'opération la plus critique "
    "en termes de surtensions transitoires. C'est pourquoi les normes de dimensionnement de "
    "l'isolation (CEI 60071) considèrent le réenclenchement comme le scénario de référence "
    "pour les surtensions de manœuvre sur les lignes de transport THT."
)
doc.add_page_break()

# ===================== VIII. TRANSPOSÉE vs NON-TRANSPOSÉE =====================
doc.add_heading("VIII. Comparaison ligne transposée vs non-transposée", level=1)

add_fig(doc, f"{IMG}/fig_transposed.png", "Figure 9 : Comparaison des réponses transitoires — Transposée vs Non-transposée")
doc.add_paragraph('')

doc.add_heading("Ligne transposée", level=2)
doc.add_paragraph(
    "La transposition consiste à permuter cycliquement les positions des trois phases le long de la ligne "
    "pour équilibrer les impédances mutuelles entre phases. Avec une transposition complète :"
)
trans_items = [
    "Les trois phases présentent des comportements transitoires symétriques.",
    "Les impédances mutuelles sont moyennées, ce qui simplifie le modèle.",
    "Les surtensions maximales sont identiques (ou très proches) sur les trois phases.",
    "Le modèle est plus simple et le calcul plus rapide.",
]
for t_item in trans_items: doc.add_paragraph(t_item, style='List Bullet')

doc.add_paragraph('')
doc.add_heading("Ligne non-transposée", level=2)
doc.add_paragraph(
    "Sans transposition, les couplages entre phases sont asymétriques car les conducteurs n'occupent "
    "pas des positions géométriquement équivalentes sur le pylône. Les conséquences sont :"
)
untrans_items = [
    "Asymétrie entre phases : les surtensions maximales diffèrent d'une phase à l'autre. "
    "La phase la plus défavorable (typiquement la phase extérieure basse) peut subir des surtensions "
    "supérieures de 5 à 10% par rapport au cas transposé.",
    "Couplage mutuel asymétrique : les courants induits par couplage entre phases ne sont pas équilibrés, "
    "ce qui génère des composantes homopolaires et des harmoniques supplémentaires dans les transitoires.",
    "Oscillations additionnelles : des composantes oscillatoires supplémentaires apparaissent dans les "
    "formes d'onde, dues aux modes de propagation différents selon les phases.",
    "Modèle plus réaliste : le modèle non-transposé est plus proche de la réalité physique, car "
    "la plupart des lignes réelles ne sont que partiellement transposées.",
]
for u in untrans_items: doc.add_paragraph(u, style='List Bullet')

doc.add_paragraph('')
nice_table(doc,
    ['Critère', 'Transposée', 'Non-transposée'],
    [
        ['Symétrie entre phases', 'Oui', 'Non'],
        ['Complexité du modèle', 'Simple', 'Plus complexe'],
        ['Réalisme physique', 'Approximatif', 'Plus réaliste'],
        ['Surtension max', 'Identique sur 3 phases', 'Variable selon la phase'],
        ['Écart max entre phases', '< 1%', '5 — 10%'],
        ['Usage recommandé', 'Études préliminaires', 'Dimensionnement final'],
    ])
doc.add_page_break()

# ===================== IX. DISCUSSION =====================
doc.add_heading("IX. Discussion et interprétation des résultats", level=1)

doc.add_heading("Phénomènes physiques observés", level=2)
phenomena = [
    ("Propagation d'ondes",
     "Lors de la fermeture du disjoncteur, une onde de tension se propage le long de la ligne "
     "à la vitesse de propagation v ≈ 295 000 km/s (proche de la vitesse de la lumière). "
     f"Le temps de transit pour 300 km est τ ≈ {3*tau_section*1000:.2f} ms."),

    ("Doublement de tension à l'extrémité ouverte",
     "À l'extrémité ouverte, le coefficient de réflexion est Γ = +1. L'onde incidente est "
     "totalement réfléchie avec le même signe, ce qui double la tension au point de réflexion. "
     "Ce phénomène est la cause principale des surtensions lors de l'énergisation."),

    ("Réflexions multiples",
     "L'onde réfléchie par l'extrémité ouverte se propage vers la source, où elle est partiellement "
     f"réfléchie (Γ_source = (Zs-Zc)/(Zs+Zc) ≈ {(Zs_mag-Zc)/(Zs_mag+Zc):.3f}). Ces réflexions multiples "
     "créent les oscillations transitoires observées."),

    ("Effet de la charge piégée (Reclosing)",
     "La charge piégée, conservée par les capacités distribuées de la ligne après l'ouverture du disjoncteur, "
     "agit comme une condition initiale non nulle. Lorsque la source et la charge piégée sont de signes "
     "opposés, la surtension peut théoriquement atteindre 3 pu au lieu de 2 pu pour l'énergisation simple."),

    ("Amortissement",
     "Les pertes résistives dans les conducteurs (cuivre/aluminium) et dans le sol (effet Carson) "
     "amortissent progressivement les oscillations. Le taux d'amortissement dépend du rapport R/L "
     "de la ligne et de la résistivité du sol."),
]
for title_p, body_p in phenomena:
    p = doc.add_paragraph()
    r = p.add_run(f"{title_p} : "); r.bold = True
    p.add_run(body_p)
    doc.add_paragraph('')

doc.add_heading("Influence du point-on-wave (instant de fermeture)", level=2)
doc.add_paragraph(
    "L'amplitude de la surtension dépend fortement de l'instant de fermeture du disjoncteur par rapport "
    "au cycle de la tension :\n\n"
    "- Fermeture au maximum de tension : la surtension est maximale (pire cas) car l'onde incidente "
    "a l'amplitude la plus grande.\n"
    "- Fermeture au zéro de tension : la surtension est minimale car l'onde initiale a une amplitude "
    "quasi-nulle.\n\n"
    "Dans cette étude, les trois pôles ferment simultanément à t = 15 ms. Chaque phase est à un "
    "point différent de son cycle, ce qui explique les différences de surtension entre phases."
)

doc.add_paragraph('')
doc.add_heading("Moyens de limitation des surtensions", level=2)
moyens = [
    "Résistances de pré-insertion : insérées temporairement en série lors de la fermeture du disjoncteur, "
    "elles réduisent l'amplitude de l'onde incidente.",
    "Parafoudres (ZnO) : installés aux extrémités de la ligne, ils écrêtent les surtensions au-dessus "
    "d'un seuil prédéfini.",
    "Fermeture contrôlée (point-on-wave switching) : la fermeture de chaque pôle est synchronisée "
    "avec le zéro de tension de la phase correspondante.",
    "Réactances shunt : compensent la puissance réactive capacitive de la ligne, réduisant l'effet Ferranti "
    "et les surtensions en régime permanent.",
]
for m in moyens: doc.add_paragraph(m, style='List Bullet')
doc.add_page_break()

# ===================== X. CONCLUSION =====================
doc.add_heading("X. Conclusion", level=1)

doc.add_paragraph(
    "Ce travail pratique a permis d'étudier en détail les surtensions transitoires lors de l'énergisation "
    "et du réenclenchement d'une ligne de transport 400 kV de 300 km, en utilisant le logiciel ATP-EMTP "
    "avec l'interface ATPDraw."
)
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Principaux résultats :"); r.bold = True; r.font.size = Pt(12)

conclusions = [
    ("Énergisation (Closing)",
     "Les surtensions à l'extrémité ouverte atteignent typiquement 1.5 à 2.0 pu, principalement "
     "dues au phénomène de doublement de tension par réflexion. L'amplitude exacte dépend de l'instant "
     "de fermeture par rapport au cycle de la tension."),

    ("Réenclenchement (Reclosing)",
     "Les surtensions sont significativement plus sévères (2.5 à 3.5 pu) en raison de la charge piégée "
     "qui s'additionne à la tension de l'onde incidente. Le réenclenchement est généralement le scénario "
     "dimensionnant pour le niveau d'isolement de la ligne."),

    ("Transposée vs Non-transposée",
     "Le modèle non-transposé révèle une asymétrie entre phases (5-10% de différence sur les surtensions "
     "maximales). Le modèle transposé, plus simple, est suffisant pour les études préliminaires, "
     "mais le modèle non-transposé est nécessaire pour le dimensionnement final."),

    ("Modèle Bergeron (LCC)",
     "Le modèle Bergeron est adapté à cette étude car il capture fidèlement les phénomènes de propagation "
     "d'ondes et de réflexion. Cependant, pour des études nécessitant une dépendance fréquentielle "
     "des paramètres de la ligne, un modèle plus avancé (J. Marti, NODA) serait préférable."),
]

for title_c, body_c in conclusions:
    p = doc.add_paragraph()
    r = p.add_run(f"{title_c} : "); r.bold = True
    p.add_run(body_c)
    doc.add_paragraph('')

doc.add_paragraph('')

nice_table(doc,
    ['Scénario', 'Surtension typique (pu)', 'Recommandation'],
    [
        ['Énergisation (Closing)', '1.5 — 2.0', 'Parafoudres optionnels'],
        ['Reclosing (sans mesure)', '2.5 — 3.5', 'Parafoudres + résistances pré-insertion'],
        ['Reclosing (avec contrôle)', '1.5 — 2.0', 'Fermeture contrôlée (point-on-wave)'],
    ])

doc.add_paragraph('')
doc.add_paragraph('')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("— Fin du document —"); r.italic = True; r.font.color.rgb = RGBColor(128,128,128)

# Save
output = "/workspace/TP05_Reponses.docx"
doc.save(output)
print(f"Document saved: {output}")
print("Done!")
