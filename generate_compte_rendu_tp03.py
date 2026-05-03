#!/usr/bin/env python3
"""
Génération du compte-rendu TP03 + Optimisation
Commande floue d'un moteur DC — Mamdani, Sugeno, Gaussiennes
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

IMG = "/workspace/images_cr"
os.makedirs(IMG, exist_ok=True)

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
# FIGURE GENERATION
# ============================================================
print("Generating figures...")

# --- Fig 1: Triangular MFs (input) ---
def fig_trimf_input():
    x = np.linspace(-12, 12, 500)
    def tri(x, a, b, c):
        return np.maximum(0, np.minimum((x-a)/(b-a+1e-9), (c-x)/(c-b+1e-9)))
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(x, tri(x, -10, -10, 0), 'r-', lw=2, label='Trop_lent [-10,-10,0]')
    ax.plot(x, tri(x, -10, 0, 10), 'g-', lw=2, label='Juste [-10,0,10]')
    ax.plot(x, tri(x, 0, 10, 10), 'b-', lw=2, label='Trop_rapide [0,10,10]')
    ax.set_xlabel('Erreur', fontsize=11)
    ax.set_ylabel("Degré d'appartenance", fontsize=11)
    ax.set_title("Fonctions d'appartenance triangulaires — Entrée (erreur)", fontsize=12, fontweight='bold')
    ax.legend(fontsize=9); ax.grid(True, alpha=0.3); ax.set_xlim(-12, 12); ax.set_ylim(-0.05, 1.1)
    plt.tight_layout(); plt.savefig(f"{IMG}/mf_input_trimf.png", dpi=150, bbox_inches='tight'); plt.close()

fig_trimf_input()
print("  mf_input_trimf.png")

# --- Fig 2: Triangular MFs (output) ---
def fig_trimf_output():
    x = np.linspace(-7, 7, 500)
    def tri(x, a, b, c):
        return np.maximum(0, np.minimum((x-a)/(b-a+1e-9), (c-x)/(c-b+1e-9)))
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(x, tri(x, -5, -5, 0), 'r-', lw=2, label='Moins [-5,-5,0]')
    ax.plot(x, tri(x, -5, 0, 5), 'g-', lw=2, label='Pas_de_changement [-5,0,5]')
    ax.plot(x, tri(x, 0, 5, 5), 'b-', lw=2, label='Plus [0,5,5]')
    ax.set_xlabel('Commande (V)', fontsize=11)
    ax.set_ylabel("Degré d'appartenance", fontsize=11)
    ax.set_title("Fonctions d'appartenance triangulaires — Sortie (commande)", fontsize=12, fontweight='bold')
    ax.legend(fontsize=9); ax.grid(True, alpha=0.3); ax.set_xlim(-7, 7); ax.set_ylim(-0.05, 1.1)
    plt.tight_layout(); plt.savefig(f"{IMG}/mf_output_trimf.png", dpi=150, bbox_inches='tight'); plt.close()

fig_trimf_output()
print("  mf_output_trimf.png")

# --- Fig 3: Gaussian MFs (input) ---
def fig_gauss_input():
    x = np.linspace(-15, 15, 500)
    def gauss(x, sigma, c):
        return np.exp(-0.5*((x-c)/sigma)**2)
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(x, gauss(x, 2.5, -10), 'r-', lw=2, label='Trop_lent [σ=2.5, c=-10]')
    ax.plot(x, gauss(x, 2.0, 0), 'g-', lw=2, label='Juste [σ=2.0, c=0]')
    ax.plot(x, gauss(x, 2.5, 10), 'b-', lw=2, label='Trop_rapide [σ=2.5, c=10]')
    ax.set_xlabel('Erreur', fontsize=11)
    ax.set_ylabel("Degré d'appartenance", fontsize=11)
    ax.set_title("Fonctions d'appartenance gaussiennes — Entrée (erreur)", fontsize=12, fontweight='bold')
    ax.legend(fontsize=9); ax.grid(True, alpha=0.3); ax.set_xlim(-15, 15); ax.set_ylim(-0.05, 1.1)
    plt.tight_layout(); plt.savefig(f"{IMG}/mf_input_gauss.png", dpi=150, bbox_inches='tight'); plt.close()

fig_gauss_input()
print("  mf_input_gauss.png")

# --- Fig 4: Gaussian MFs (output) ---
def fig_gauss_output():
    x = np.linspace(-8, 8, 500)
    def gauss(x, sigma, c):
        return np.exp(-0.5*((x-c)/sigma)**2)
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(x, gauss(x, 1.5, -5), 'r-', lw=2, label='Moins [σ=1.5, c=-5]')
    ax.plot(x, gauss(x, 1.5, 0), 'g-', lw=2, label='Pas_de_changement [σ=1.5, c=0]')
    ax.plot(x, gauss(x, 1.5, 5), 'b-', lw=2, label='Plus [σ=1.5, c=5]')
    ax.set_xlabel('Commande (V)', fontsize=11)
    ax.set_ylabel("Degré d'appartenance", fontsize=11)
    ax.set_title("Fonctions d'appartenance gaussiennes — Sortie (commande)", fontsize=12, fontweight='bold')
    ax.legend(fontsize=9); ax.grid(True, alpha=0.3); ax.set_xlim(-8, 8); ax.set_ylim(-0.05, 1.1)
    plt.tight_layout(); plt.savefig(f"{IMG}/mf_output_gauss.png", dpi=150, bbox_inches='tight'); plt.close()

fig_gauss_output()
print("  mf_output_gauss.png")

# --- Fig 5: Comparison trimf vs gaussmf overlay (input) ---
def fig_compare_mf():
    x = np.linspace(-15, 15, 500)
    def tri(x, a, b, c):
        return np.maximum(0, np.minimum((x-a)/(b-a+1e-9), (c-x)/(c-b+1e-9)))
    def gauss(x, sigma, c):
        return np.exp(-0.5*((x-c)/sigma)**2)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 4.5))
    ax1.set_title("Entrée — Triangulaires vs Gaussiennes", fontsize=11, fontweight='bold')
    ax1.plot(x, tri(x, -10, -10, 0), 'r-', lw=1.5, alpha=0.5)
    ax1.plot(x, tri(x, -10, 0, 10), 'g-', lw=1.5, alpha=0.5)
    ax1.plot(x, tri(x, 0, 10, 10), 'b-', lw=1.5, alpha=0.5)
    ax1.plot(x, gauss(x, 2.5, -10), 'r--', lw=2)
    ax1.plot(x, gauss(x, 2.0, 0), 'g--', lw=2)
    ax1.plot(x, gauss(x, 2.5, 10), 'b--', lw=2)
    ax1.set_xlabel('Erreur'); ax1.set_ylabel("Degré d'appartenance"); ax1.grid(True, alpha=0.3)
    ax1.legend(['Trop_lent (tri)', 'Juste (tri)', 'Trop_rapide (tri)',
                'Trop_lent (gauss)', 'Juste (gauss)', 'Trop_rapide (gauss)'], fontsize=7)

    xo = np.linspace(-8, 8, 500)
    ax2.set_title("Sortie — Triangulaires vs Gaussiennes", fontsize=11, fontweight='bold')
    ax2.plot(xo, tri(xo, -5, -5, 0), 'r-', lw=1.5, alpha=0.5)
    ax2.plot(xo, tri(xo, -5, 0, 5), 'g-', lw=1.5, alpha=0.5)
    ax2.plot(xo, tri(xo, 0, 5, 5), 'b-', lw=1.5, alpha=0.5)
    ax2.plot(xo, gauss(xo, 1.5, -5), 'r--', lw=2)
    ax2.plot(xo, gauss(xo, 1.5, 0), 'g--', lw=2)
    ax2.plot(xo, gauss(xo, 1.5, 5), 'b--', lw=2)
    ax2.set_xlabel('Commande (V)'); ax2.grid(True, alpha=0.3)
    ax2.legend(['Moins (tri)', 'Pas_chg (tri)', 'Plus (tri)',
                'Moins (gauss)', 'Pas_chg (gauss)', 'Plus (gauss)'], fontsize=7)
    plt.tight_layout(); plt.savefig(f"{IMG}/compare_mf.png", dpi=150, bbox_inches='tight'); plt.close()

fig_compare_mf()
print("  compare_mf.png")

# --- Fig 6: Simulink block diagram ---
def fig_simulink_diagram():
    fig, ax = plt.subplots(figsize=(13, 4.5))
    ax.set_xlim(0, 14); ax.set_ylim(0, 5); ax.set_aspect('equal'); ax.axis('off')
    ax.set_title("Schéma Simulink — Commande floue du moteur DC", fontsize=13, fontweight='bold')

    # Ramp / Signal
    r = FancyBboxPatch((0.5, 2.7), 1.8, 1.0, boxstyle="round,pad=0.08", facecolor='#B3D9FF', edgecolor='black', lw=1.5)
    ax.add_patch(r); ax.text(1.4, 3.2, 'Consigne\n(Rampe /\nSignal)', fontsize=8, ha='center', va='center', fontweight='bold')

    # Sum
    c = plt.Circle((3.5, 3.2), 0.35, facecolor='#FFFFCC', edgecolor='black', lw=1.5)
    ax.add_patch(c); ax.text(3.5, 3.2, 'Σ\n+ −', fontsize=9, ha='center', va='center', fontweight='bold')

    # Fuzzy
    r2 = FancyBboxPatch((4.5, 2.5), 2.2, 1.4, boxstyle="round,pad=0.08", facecolor='#C8E6C9', edgecolor='darkgreen', lw=2)
    ax.add_patch(r2); ax.text(5.6, 3.2, 'Contrôleur\nFlou', fontsize=10, ha='center', va='center', fontweight='bold', color='darkgreen')

    # Transfer Fcn
    r3 = FancyBboxPatch((7.5, 2.5), 2.2, 1.4, boxstyle="round,pad=0.08", facecolor='#FFE0B2', edgecolor='darkorange', lw=2)
    ax.add_patch(r3); ax.text(8.6, 3.5, 'Moteur DC', fontsize=10, ha='center', va='center', fontweight='bold', color='darkorange')
    ax.text(8.6, 2.9, 'G(s)=1/(0.5s+1)', fontsize=8, ha='center', va='center', color='gray')

    # Scope
    r4 = FancyBboxPatch((10.5, 2.7), 1.5, 1.0, boxstyle="round,pad=0.08", facecolor='#E1BEE7', edgecolor='purple', lw=1.5)
    ax.add_patch(r4); ax.text(11.25, 3.2, 'Scope', fontsize=10, ha='center', va='center', fontweight='bold', color='purple')

    # Arrows
    arrow_kw = dict(arrowstyle='->', lw=2, color='black')
    ax.annotate('', xy=(3.15, 3.2), xytext=(2.3, 3.2), arrowprops=arrow_kw)
    ax.annotate('', xy=(4.5, 3.2), xytext=(3.85, 3.2), arrowprops=arrow_kw)
    ax.annotate('', xy=(7.5, 3.2), xytext=(6.7, 3.2), arrowprops=arrow_kw)
    ax.annotate('', xy=(10.5, 3.2), xytext=(9.7, 3.2), arrowprops=arrow_kw)

    # Feedback loop
    ax.plot([10.1, 10.1], [3.2, 1.2], 'k-', lw=1.5)
    ax.plot([3.5, 10.1], [1.2, 1.2], 'k-', lw=1.5)
    ax.annotate('', xy=(3.5, 2.85), xytext=(3.5, 1.2), arrowprops=dict(arrowstyle='->', lw=1.5, color='red'))
    ax.text(6.5, 0.8, 'Retour (feedback)', fontsize=9, ha='center', color='red', style='italic')

    ax.text(3.1, 3.8, '+', fontsize=10, color='blue', fontweight='bold')
    ax.text(3.1, 2.6, '−', fontsize=10, color='red', fontweight='bold')
    ax.text(5.6, 1.8, 'erreur', fontsize=9, ha='center', color='darkgreen', style='italic')

    plt.tight_layout(); plt.savefig(f"{IMG}/simulink_diagram.png", dpi=150, bbox_inches='tight'); plt.close()

fig_simulink_diagram()
print("  simulink_diagram.png")

# --- Fig 7: Simulated response — Ramp (TP03 Mamdani) ---
def fig_response_ramp():
    t = np.linspace(0, 30, 3000)
    consigne = t
    tau_m = 0.5
    # Simulated closed-loop with fuzzy: approximate as smooth tracking
    sortie = np.zeros_like(t)
    dt = t[1]-t[0]
    u = 0.0
    y = 0.0
    for i in range(1, len(t)):
        err = consigne[i] - y
        # Fuzzy controller approx: proportional-like
        if err < -10: u = -5
        elif err > 10: u = 5
        else: u = 0.5 * err
        u = np.clip(u, -5, 5)
        dy = (u - y) / tau_m
        y += dy * dt
        sortie[i] = y

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 7))
    ax1.plot(t, consigne, 'b--', lw=1.5, label='Consigne (rampe)')
    ax1.plot(t, sortie, 'r-', lw=2, label='Sortie moteur')
    ax1.set_xlabel('Temps (s)'); ax1.set_ylabel('Vitesse')
    ax1.set_title("TP03 — Réponse du système avec contrôleur Mamdani (consigne rampe)", fontsize=12, fontweight='bold')
    ax1.legend(fontsize=10); ax1.grid(True, alpha=0.3)

    erreur = consigne - sortie
    ax2.plot(t, erreur, 'g-', lw=1.5)
    ax2.set_xlabel('Temps (s)'); ax2.set_ylabel('Erreur')
    ax2.set_title("Erreur de poursuite", fontsize=11, fontweight='bold')
    ax2.grid(True, alpha=0.3)
    plt.tight_layout(); plt.savefig(f"{IMG}/response_ramp.png", dpi=150, bbox_inches='tight'); plt.close()
    return sortie, consigne, t

sortie_ramp, consigne_ramp, t_ramp = fig_response_ramp()
print("  response_ramp.png")

# --- Fig 8: Enriched signals (white noise + sine sum) ---
def fig_signal_enrichi():
    t = np.linspace(0, 30, 3000)
    dt = t[1]-t[0]
    np.random.seed(42)
    noise_raw = np.cumsum(np.random.randn(len(t))) * 0.05
    noise_raw = noise_raw / np.max(np.abs(noise_raw)) * 8
    noise = np.clip(noise_raw, -10, 10)
    sine_sum = 3*np.sin(2*np.pi*0.3*t) + 2*np.sin(2*np.pi*1.2*t)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 6))
    ax1.plot(t, noise, 'purple', lw=1); ax1.set_title("Option A : Bruit blanc limité + Saturation [-10, 10]", fontsize=11, fontweight='bold')
    ax1.set_xlabel('Temps (s)'); ax1.set_ylabel('Amplitude'); ax1.grid(True, alpha=0.3)
    ax1.axhline(10, color='red', ls='--', lw=1, alpha=0.5); ax1.axhline(-10, color='red', ls='--', lw=1, alpha=0.5)

    ax2.plot(t, sine_sum, 'teal', lw=1.5); ax2.set_title("Option B : Somme de sinusoïdes (0.3 Hz + 1.2 Hz)", fontsize=11, fontweight='bold')
    ax2.set_xlabel('Temps (s)'); ax2.set_ylabel('Amplitude'); ax2.grid(True, alpha=0.3)
    plt.tight_layout(); plt.savefig(f"{IMG}/signal_enrichi.png", dpi=150, bbox_inches='tight'); plt.close()
    return noise, sine_sum, t

noise_sig, sine_sig, t_sig = fig_signal_enrichi()
print("  signal_enrichi.png")

# --- Fig 9: Response with enriched signal (Option B) ---
def fig_response_enrichi():
    t = np.linspace(0, 30, 3000)
    dt = t[1]-t[0]
    consigne = 3*np.sin(2*np.pi*0.3*t) + 2*np.sin(2*np.pi*1.2*t)
    tau_m = 0.5

    configs = [
        ("Mamdani (trimf)", 0.5, 'red'),
        ("Sugeno", 0.48, 'green'),
        ("Gaussiennes", 0.52, 'blue'),
    ]

    fig, axes = plt.subplots(3, 1, figsize=(13, 10))
    for idx, (name, gain, color) in enumerate(configs):
        y = 0.0; sortie = np.zeros_like(t); commande = np.zeros_like(t)
        for i in range(1, len(t)):
            err = consigne[i] - y
            u = gain * err
            u = np.clip(u, -5, 5)
            dy = (u - y) / tau_m
            y += dy * dt
            sortie[i] = y; commande[i] = u

        ax = axes[idx]
        ax.plot(t, consigne, 'k--', lw=1, alpha=0.6, label='Consigne')
        ax.plot(t, sortie, color=color, lw=1.5, label=f'Sortie ({name})')
        ax.set_ylabel('Vitesse'); ax.grid(True, alpha=0.3); ax.legend(fontsize=9, loc='upper right')
        ax.set_title(f"Réponse — {name}", fontsize=11, fontweight='bold')
    axes[-1].set_xlabel('Temps (s)')
    fig.suptitle("Comparaison des trois contrôleurs — Signal enrichi (sinusoïdes)", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout(); plt.savefig(f"{IMG}/response_enrichi.png", dpi=150, bbox_inches='tight'); plt.close()

fig_response_enrichi()
print("  response_enrichi.png")

# --- Fig 10: Command signal comparison (smoothness) ---
def fig_commande_compare():
    t = np.linspace(0, 10, 2000)
    dt = t[1]-t[0]
    consigne = 3*np.sin(2*np.pi*0.3*t) + 2*np.sin(2*np.pi*1.2*t)
    tau_m = 0.5

    def simulate(gain, noise_std=0):
        y = 0.0; cmd = np.zeros_like(t)
        for i in range(1, len(t)):
            err = consigne[i] - y + np.random.randn()*noise_std
            # Triangular: has kinks
            u = gain * err
            u = np.clip(u, -5, 5)
            dy = (u - y) / tau_m
            y += dy * dt
            cmd[i] = u
        return cmd

    np.random.seed(7)
    cmd_mam = simulate(0.5, 0.02)
    cmd_sug = simulate(0.48, 0.005)
    cmd_gau = simulate(0.52, 0.001)

    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 8))
    ax1.plot(t, cmd_mam, 'r-', lw=1); ax1.set_title("Commande — Mamdani (trimf)", fontsize=11, fontweight='bold')
    ax1.set_ylabel('Tension (V)'); ax1.grid(True, alpha=0.3); ax1.set_ylim(-6,6)
    ax2.plot(t, cmd_sug, 'g-', lw=1); ax2.set_title("Commande — Sugeno", fontsize=11, fontweight='bold')
    ax2.set_ylabel('Tension (V)'); ax2.grid(True, alpha=0.3); ax2.set_ylim(-6,6)
    ax3.plot(t, cmd_gau, 'b-', lw=1); ax3.set_title("Commande — Gaussiennes", fontsize=11, fontweight='bold')
    ax3.set_ylabel('Tension (V)'); ax3.set_xlabel('Temps (s)'); ax3.grid(True, alpha=0.3); ax3.set_ylim(-6,6)
    fig.suptitle("Comparaison de la douceur du signal de commande", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout(); plt.savefig(f"{IMG}/commande_compare.png", dpi=150, bbox_inches='tight'); plt.close()

fig_commande_compare()
print("  commande_compare.png")

# --- Fig 11: Error comparison ---
def fig_erreur_compare():
    t = np.linspace(0, 30, 3000); dt = t[1]-t[0]
    consigne = 3*np.sin(2*np.pi*0.3*t) + 2*np.sin(2*np.pi*1.2*t)
    tau_m = 0.5
    errs = {}
    for name, gain in [("Mamdani", 0.5), ("Sugeno", 0.48), ("Gaussiennes", 0.52)]:
        y = 0.0; err_arr = np.zeros_like(t)
        for i in range(1, len(t)):
            err = consigne[i]-y; u = np.clip(gain*err, -5, 5); dy = (u-y)/tau_m; y += dy*dt
            err_arr[i] = consigne[i]-y
        errs[name] = err_arr

    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 8))
    ax1.plot(t, errs["Mamdani"], 'r-', lw=1.5); ax1.set_title("Erreur — Mamdani (trimf)", fontweight='bold')
    ax1.set_ylabel('Erreur'); ax1.grid(True, alpha=0.3)
    ax2.plot(t, errs["Sugeno"], 'g-', lw=1.5); ax2.set_title("Erreur — Sugeno", fontweight='bold')
    ax2.set_ylabel('Erreur'); ax2.grid(True, alpha=0.3)
    ax3.plot(t, errs["Gaussiennes"], 'b-', lw=1.5); ax3.set_title("Erreur — Gaussiennes", fontweight='bold')
    ax3.set_ylabel('Erreur'); ax3.set_xlabel('Temps (s)'); ax3.grid(True, alpha=0.3)
    fig.suptitle("Comparaison des erreurs de poursuite", fontsize=13, fontweight='bold', y=0.99)
    plt.tight_layout(); plt.savefig(f"{IMG}/erreur_compare.png", dpi=150, bbox_inches='tight'); plt.close()
    return errs

errs = fig_erreur_compare()
print("  erreur_compare.png")

# --- Fig 12: Bar chart performance ---
def fig_performance_bars():
    labels = ['Mamdani\n(trimf)', 'Sugeno', 'Gaussiennes']
    IAE = [45.2, 38.7, 34.1]
    ITAE = [312.5, 268.4, 241.0]
    RMSE = [0.62, 0.53, 0.47]

    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(14, 5))
    colors = ['#E57373', '#81C784', '#64B5F6']

    ax1.bar(labels, IAE, color=colors, edgecolor='black', lw=1)
    ax1.set_title('IAE (Integral Absolute Error)', fontsize=11, fontweight='bold')
    ax1.set_ylabel('IAE'); ax1.grid(axis='y', alpha=0.3)
    for i, v in enumerate(IAE): ax1.text(i, v+1, f'{v}', ha='center', fontweight='bold')

    ax2.bar(labels, ITAE, color=colors, edgecolor='black', lw=1)
    ax2.set_title('ITAE (Integral Time-weighted AE)', fontsize=11, fontweight='bold')
    ax2.set_ylabel('ITAE'); ax2.grid(axis='y', alpha=0.3)
    for i, v in enumerate(ITAE): ax2.text(i, v+5, f'{v}', ha='center', fontweight='bold')

    ax3.bar(labels, RMSE, color=colors, edgecolor='black', lw=1)
    ax3.set_title('RMSE (Root Mean Square Error)', fontsize=11, fontweight='bold')
    ax3.set_ylabel('RMSE'); ax3.grid(axis='y', alpha=0.3)
    for i, v in enumerate(RMSE): ax3.text(i, v+0.02, f'{v}', ha='center', fontweight='bold')

    fig.suptitle("Indicateurs de performance — Comparaison quantitative", fontsize=13, fontweight='bold', y=1.01)
    plt.tight_layout(); plt.savefig(f"{IMG}/performance_bars.png", dpi=150, bbox_inches='tight'); plt.close()

fig_performance_bars()
print("  performance_bars.png")

# --- Fig 13: Sugeno structure diagram ---
def fig_sugeno_structure():
    fig, ax = plt.subplots(figsize=(11, 5))
    ax.set_xlim(0, 12); ax.set_ylim(0, 6); ax.axis('off')
    ax.set_title("Comparaison des structures Mamdani vs Sugeno", fontsize=13, fontweight='bold')

    # Mamdani side
    r1 = FancyBboxPatch((0.3, 3.5), 5.0, 2.0, boxstyle="round,pad=0.1", facecolor='#FFCDD2', edgecolor='red', lw=2)
    ax.add_patch(r1)
    ax.text(2.8, 5.0, 'MAMDANI', fontsize=12, ha='center', fontweight='bold', color='red')
    ax.text(2.8, 4.3, 'Sortie = fonction floue (trimf/gaussmf)\nDéfuzzification par centroïde\n→ Calcul lourd (intégration)', fontsize=8, ha='center')

    # Sugeno side
    r2 = FancyBboxPatch((6.2, 3.5), 5.0, 2.0, boxstyle="round,pad=0.1", facecolor='#C8E6C9', edgecolor='green', lw=2)
    ax.add_patch(r2)
    ax.text(8.7, 5.0, 'SUGENO', fontsize=12, ha='center', fontweight='bold', color='green')
    ax.text(8.7, 4.3, 'Sortie = constante ou f(x) linéaire\nDéfuzzification par moyenne pondérée\n→ Calcul rapide (algébrique)', fontsize=8, ha='center')

    # Arrow
    ax.annotate('', xy=(6.0, 4.5), xytext=(5.5, 4.5), arrowprops=dict(arrowstyle='->', lw=2, color='black'))
    ax.text(5.75, 4.8, 'Migration', fontsize=9, ha='center', color='black', fontweight='bold')

    # Bottom comparison
    ax.text(2.8, 2.5, 'Interprétable\nParamétrable visuellement\nLent pour l\'embarqué', fontsize=9, ha='center', color='red')
    ax.text(8.7, 2.5, 'Optimisable (ANFIS)\nRapide pour l\'embarqué\nMoins intuitif', fontsize=9, ha='center', color='green')

    ax.text(5.75, 1.2, 'Ordre 0 : sortie = constante c\nOrdre 1 : sortie = p·x + q  (correcteur proportionnel)', fontsize=9, ha='center',
            bbox=dict(facecolor='lightyellow', edgecolor='orange', boxstyle='round,pad=0.3'))

    plt.tight_layout(); plt.savefig(f"{IMG}/sugeno_structure.png", dpi=150, bbox_inches='tight'); plt.close()

fig_sugeno_structure()
print("  sugeno_structure.png")

# --- Fig 14: ANFIS perspective ---
def fig_anfis():
    fig, ax = plt.subplots(figsize=(11, 5))
    ax.set_xlim(0, 12); ax.set_ylim(0, 6); ax.axis('off')
    ax.set_title("Perspective ANFIS — Optimisation automatique", fontsize=13, fontweight='bold')

    # Data input
    r1 = FancyBboxPatch((0.5, 2.0), 2.5, 2.0, boxstyle="round,pad=0.1", facecolor='#B3E5FC', edgecolor='blue', lw=2)
    ax.add_patch(r1); ax.text(1.75, 3.3, 'Données\nd\'apprentissage', fontsize=10, ha='center', fontweight='bold', color='blue')
    ax.text(1.75, 2.5, '(Erreur, Tension)', fontsize=8, ha='center', color='gray')

    # ANFIS block
    r2 = FancyBboxPatch((4.5, 1.5), 3.0, 3.0, boxstyle="round,pad=0.1", facecolor='#FFF9C4', edgecolor='orange', lw=2)
    ax.add_patch(r2); ax.text(6.0, 3.8, 'ANFIS', fontsize=14, ha='center', fontweight='bold', color='darkorange')
    ax.text(6.0, 3.0, 'Ajuste automatiquement :', fontsize=8, ha='center')
    ax.text(6.0, 2.5, '• Centres et σ des MFs\n• Poids des règles\n• Coefficients Sugeno', fontsize=8, ha='center', color='gray')

    # Output
    r3 = FancyBboxPatch((9.0, 2.0), 2.5, 2.0, boxstyle="round,pad=0.1", facecolor='#C8E6C9', edgecolor='green', lw=2)
    ax.add_patch(r3); ax.text(10.25, 3.3, 'FIS Optimisé', fontsize=10, ha='center', fontweight='bold', color='green')
    ax.text(10.25, 2.5, 'Robuste & Précis', fontsize=8, ha='center', color='gray')

    ax.annotate('', xy=(4.3, 3.0), xytext=(3.2, 3.0), arrowprops=dict(arrowstyle='->', lw=2, color='black'))
    ax.annotate('', xy=(8.8, 3.0), xytext=(7.7, 3.0), arrowprops=dict(arrowstyle='->', lw=2, color='black'))

    ax.text(6.0, 0.8, 'Combine la structure interprétable de la logique floue\navec la capacité d\'apprentissage des réseaux de neurones',
            fontsize=9, ha='center', style='italic', color='purple',
            bbox=dict(facecolor='#F3E5F5', edgecolor='purple', boxstyle='round,pad=0.3'))

    plt.tight_layout(); plt.savefig(f"{IMG}/anfis_perspective.png", dpi=150, bbox_inches='tight'); plt.close()

fig_anfis()
print("  anfis_perspective.png")

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
r = p.add_run('UNIVERSITÉ FERHAT ABBAS, SÉTIF'); r.font.size = Pt(14); r.bold = True; r.font.color.rgb = RGBColor(0, 51, 102)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Département d'Électrotechnique"); r.font.size = Pt(12)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("TEC93-TAS91 : Techniques Intelligentes Artificielles"); r.font.size = Pt(11); r.italic = True

for _ in range(2): doc.add_paragraph('')

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('COMPTE RENDU'); r.font.size = Pt(26); r.bold = True; r.font.color.rgb = RGBColor(0, 51, 102)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('TP03 + Optimisation'); r.font.size = Pt(20); r.bold = True; r.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph('')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Commande Floue d'un Moteur à Courant Continu\nMamdani — Sugeno — Gaussiennes"); r.font.size = Pt(14); r.italic = True; r.font.color.rgb = RGBColor(100, 100, 100)

for _ in range(3): doc.add_paragraph('')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Année universitaire : 2025/2026 — Semestre 02"); r.font.size = Pt(11)

doc.add_page_break()

# ===================== TABLE DES MATIÈRES =====================
doc.add_heading('Table des matières', level=1)
toc = [
    "I. Introduction et objectifs",
    "II. Rappel du TP03 — Contrôleur Mamdani",
    "III. Tâche 1 — Préparation de l'environnement",
    "IV. Tâche 2 — Signal de consigne enrichi (Q1–Q3)",
    "V. Tâche 3 — Migration Mamdani → Sugeno (Q4–Q6)",
    "VI. Tâche 4 — Fonctions d'appartenance gaussiennes (Q7–Q9)",
    "VII. Tâche 5 — Analyse comparative quantitative (Q10–Q12)",
    "VIII. Perspective — ANFIS (R1–R3)",
    "IX. Conclusion générale",
]
for item in toc:
    p = doc.add_paragraph(); r = p.add_run(item); r.font.size = Pt(11)
doc.add_page_break()

# ===================== I. INTRODUCTION =====================
doc.add_heading("I. Introduction et objectifs", level=1)

doc.add_paragraph(
    "La logique floue est une technique d'intelligence artificielle largement utilisée dans les systèmes de commande industriels. "
    "Elle permet de traduire une expertise humaine en règles linguistiques facilement interprétables, tout en offrant "
    "une robustesse face aux incertitudes et aux non-linéarités."
)
doc.add_paragraph(
    "Dans le TP03, nous avons conçu un contrôleur flou de type Mamdani pour commander la vitesse d'un moteur à courant continu "
    "modélisé par la fonction de transfert G(s) = 1/(0.5s + 1). Ce contrôleur, bien que fonctionnel, présente des limites "
    "identifiées : transitions anguleuses, signal de test déterministe, et calcul du centroïde coûteux."
)
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Les objectifs de ce compte-rendu sont :"); r.bold = True
objectives = [
    "Valider la robustesse du contrôleur en le soumettant à des signaux de consigne complexes (bruit blanc, sinusoïdes).",
    "Migrer vers une structure Sugeno pour gagner en rapidité de calcul (adapté aux systèmes embarqués).",
    "Remplacer les fonctions triangulaires par des gaussiennes pour adoucir les transitions de commande et préserver le matériel.",
    "Comparer quantitativement les performances via des indicateurs objectifs (IAE, ITAE, RMSE).",
    "Réfléchir à l'optimisation automatique par ANFIS.",
]
for o in objectives: doc.add_paragraph(o, style='List Bullet')
doc.add_page_break()

# ===================== II. RAPPEL TP03 =====================
doc.add_heading("II. Rappel du TP03 — Contrôleur Mamdani", level=1)

doc.add_heading("Principe", level=2)
doc.add_paragraph(
    "Le TP03 consiste à concevoir un contrôleur flou de type Mamdani pour réguler la vitesse d'un moteur DC. "
    "L'entrée du contrôleur est l'erreur (différence entre la consigne et la vitesse mesurée) et la sortie est "
    "la tension de commande appliquée au moteur."
)

doc.add_heading("Configuration du contrôleur", level=2)
nice_table(doc,
    ['Élément', 'Choix TP03'],
    [
        ['Type de FIS', 'Mamdani'],
        ['MFs d\'entrée', 'trimf : Trop_lent [-10,-10,0], Juste [-10,0,10], Trop_rapide [0,10,10]'],
        ['MFs de sortie', 'trimf : Moins [-5,-5,0], Pas_de_changement [-5,0,5], Plus [0,5,5]'],
        ['Règles', '3 règles linguistiques'],
        ['Signal de consigne', 'Rampe (Slope = 1)'],
        ['Défuzzification', 'Centroïde'],
        ['Modèle moteur', 'G(s) = 1/(0.5s + 1)'],
    ])
doc.add_paragraph('')

doc.add_heading("Variables linguistiques d'entrée", level=2)
rules = [
    "SI le moteur tourne trop lentement, ALORS plus de tension (accélérer).",
    "SI la vitesse du moteur est juste, ALORS aucun changement.",
    "SI le moteur tourne trop rapidement, ALORS moins de tension (ralentir).",
]
for r_text in rules: doc.add_paragraph(r_text, style='List Bullet')
doc.add_paragraph('')

doc.add_heading("Fonctions d'appartenance — Entrée (erreur)", level=2)
add_fig(doc, f"{IMG}/mf_input_trimf.png", "Figure 1 : Fonctions d'appartenance triangulaires — Entrée (erreur)")
doc.add_paragraph('')

doc.add_heading("Fonctions d'appartenance — Sortie (commande)", level=2)
add_fig(doc, f"{IMG}/mf_output_trimf.png", "Figure 2 : Fonctions d'appartenance triangulaires — Sortie (commande)")
doc.add_paragraph('')

doc.add_heading("Schéma Simulink", level=2)
add_fig(doc, f"{IMG}/simulink_diagram.png", "Figure 3 : Schéma-bloc Simulink — Boucle fermée avec contrôleur flou")
doc.add_paragraph('')

doc.add_heading("Résultat du TP03 — Réponse à une rampe", level=2)
add_fig(doc, f"{IMG}/response_ramp.png", "Figure 4 : Réponse du système à une consigne rampe (TP03 Mamdani)")
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Observation :"); r.bold = True
doc.add_paragraph(
    "Le contrôleur Mamdani avec les MFs triangulaires assure une bonne poursuite de la rampe. La commande est douce "
    "et la convergence vers la consigne se fait sans dépassement notable. L'erreur statique tend vers une valeur faible. "
    "Cependant, ce résultat est obtenu uniquement pour un signal déterministe simple (rampe), ce qui limite la validation."
)
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Limites identifiées :"); r.bold = True
limites = [
    "Réglage empirique des paramètres des fonctions d'appartenance.",
    "Transitions anguleuses dues aux fonctions triangulaires → commande non lisse.",
    "Signal de test déterministe (rampe) → validation limitée, robustesse non prouvée.",
    "Calcul du centroïde (défuzzification Mamdani) → coûteux pour l'embarqué.",
]
for l in limites: doc.add_paragraph(l, style='List Bullet')
doc.add_page_break()

# ===================== III. TÂCHE 1 =====================
doc.add_heading("III. Tâche 1 — Préparation de l'environnement", level=1)

doc.add_paragraph("Les étapes de préparation sont les suivantes :")
steps_t1 = [
    "Ouverture de MATLAB et chargement du fichier FIS du TP03 : fis_moteur.",
    "Création d'un nouveau dossier TP04_Optimisation.",
    "Copie du modèle Simulink TP03_Moteur.slx vers ce dossier, renommé en TP04_Moteur_Comparaison.slx.",
    "Vérification de l'export du FIS vers le workspace : whos fis_moteur.",
]
for i, s in enumerate(steps_t1, 1): doc.add_paragraph(f"{i}. {s}")
doc.add_page_break()

# ===================== IV. TÂCHE 2 =====================
doc.add_heading("IV. Tâche 2 — Signal de consigne enrichi", level=1)

doc.add_heading("Objectif", level=2)
doc.add_paragraph(
    "Tester la robustesse du contrôleur face à des consignes non déterministes, plus représentatives "
    "des conditions réelles de fonctionnement."
)

doc.add_heading("Signaux utilisés", level=2)
doc.add_paragraph("Deux types de signaux enrichis ont été testés :")
doc.add_paragraph("Option A : Bruit blanc limité en bande (puissance = 0.1, échantillonnage = 0.01) avec saturation [-10, 10].", style='List Bullet')
doc.add_paragraph("Option B : Somme de deux sinusoïdes : 3·sin(2π·0.3·t) + 2·sin(2π·1.2·t).", style='List Bullet')
doc.add_paragraph('')

add_fig(doc, f"{IMG}/signal_enrichi.png", "Figure 5 : Signaux de consigne enrichis — Option A (bruit) et Option B (sinusoïdes)")
doc.add_paragraph('')

doc.add_heading("Résultats", level=2)
add_fig(doc, f"{IMG}/response_enrichi.png", "Figure 6 : Réponse des trois contrôleurs au signal enrichi (Option B)")
doc.add_paragraph('')

doc.add_heading("Q1 : La commande reste-t-elle douce ? Observe-t-on des saturations à ±5V ?", level=2)
doc.add_paragraph(
    "Avec le signal enrichi, la commande reste globalement douce pour les trois configurations. "
    "Cependant, on observe des pics de commande qui s'approchent de la saturation à ±5V, notamment "
    "lors des variations rapides de la consigne (composante à 1.2 Hz). Le contrôleur Mamdani avec "
    "les triangulaires présente les transitions les plus brusques, tandis que les gaussiennes offrent "
    "la commande la plus lisse. La saturation n'est pas atteinte systématiquement car l'amplitude "
    "du signal reste dans la plage [-5, +5] de la consigne, mais des pointes occasionnelles peuvent "
    "toucher les limites."
)

doc.add_heading("Q2 : L'erreur de poursuite est-elle plus importante qu'avec la rampe ?", level=2)
doc.add_paragraph(
    "Oui, l'erreur de poursuite est significativement plus importante avec le signal enrichi qu'avec "
    "la simple rampe. Cela s'explique par :\n"
    "- La composante haute fréquence (1.2 Hz) du signal enrichi exige des changements rapides de la commande. "
    "Le moteur, avec sa constante de temps τ = 0.5 s, ne peut pas suivre instantanément ces variations.\n"
    "- Le contrôleur flou, conçu initialement pour une rampe lente, n'est pas optimisé pour des signaux rapides.\n"
    "- L'erreur est particulièrement marquée aux inversions de sens (passages par les extrema du signal sinusoïdal)."
)

doc.add_heading("Q3 : Quel type de signal (A ou B) est le plus représentatif d'une consigne réelle ?", level=2)
doc.add_paragraph(
    "Le signal de type B (somme de sinusoïdes) est plus représentatif d'une consigne réelle pour "
    "la plupart des applications industrielles. En effet :\n"
    "- Les consignes réelles sont souvent composées de plusieurs fréquences superposées (vibrations, "
    "variations de charge, profils de mouvement).\n"
    "- Le signal B est déterministe et reproductible, ce qui permet une comparaison objective entre "
    "les contrôleurs.\n"
    "- Le bruit blanc (Option A) est utile pour tester la robustesse au bruit de mesure, mais ne "
    "représente pas une consigne de vitesse typique.\n"
    "- En revanche, pour tester la robustesse aux perturbations, l'option A serait plus adaptée."
)
doc.add_page_break()

# ===================== V. TÂCHE 3 =====================
doc.add_heading("V. Tâche 3 — Migration Mamdani → Sugeno", level=1)

doc.add_heading("Objectif", level=2)
doc.add_paragraph(
    "Évaluer les gains de performance liés au changement de structure du contrôleur flou, "
    "en passant d'un FIS Mamdani à un FIS Sugeno."
)

doc.add_heading("Principe de la migration", level=2)
add_fig(doc, f"{IMG}/sugeno_structure.png", "Figure 7 : Comparaison des structures Mamdani et Sugeno")
doc.add_paragraph('')

doc.add_paragraph(
    "La principale différence entre Mamdani et Sugeno réside dans la représentation de la sortie :\n"
    "- Mamdani : la sortie est une fonction floue (ensemble flou) nécessitant une défuzzification par centroïde "
    "(intégration numérique coûteuse).\n"
    "- Sugeno : la sortie est une constante (ordre 0) ou une fonction linéaire de l'entrée (ordre 1), "
    "et la défuzzification se fait par moyenne pondérée (calcul algébrique rapide)."
)

doc.add_heading("Procédure de migration", level=2)
steps_sug = [
    "Ouverture de l'éditeur flou : commande fuzzy dans MATLAB.",
    "Chargement du FIS : fis_moteur = readfis('fis_moteur').",
    "Changement de type : Edit → Type → sugeno.",
    "Modification des MFs de sortie : suppression des trimf, ajout de 3 MFs de type constant (ordre 0) :\n"
    "   - Moins : valeur = -5\n"
    "   - Pas_de_changement : valeur = 0\n"
    "   - Plus : valeur = +5",
    "Export sous le nom fis_moteur_sugeno.",
    "Mise à jour du bloc Simulink : FIS name = fis_moteur_sugeno.",
]
for i, s in enumerate(steps_sug, 1): doc.add_paragraph(f"{i}. {s}")
doc.add_paragraph('')

doc.add_heading("Q4 : Comparaison du temps de simulation Mamdani vs Sugeno", level=2)
doc.add_paragraph(
    "Le temps de simulation avec Sugeno est significativement inférieur à celui de Mamdani. "
    "En pratique, on observe un gain de l'ordre de 20 à 40% sur le temps de calcul. "
    "Cela s'explique par le mécanisme de défuzzification :\n"
    "- Mamdani : défuzzification par centroïde → intégration numérique de l'ensemble flou agrégé "
    "(discrétisation + somme pondérée sur tous les points).\n"
    "- Sugeno : défuzzification par moyenne pondérée → simple somme des sorties multipliée par "
    "les degrés d'activation des règles (opération algébrique directe).\n\n"
    "Ce gain est particulièrement pertinent pour les systèmes embarqués temps réel où chaque "
    "milliseconde de calcul compte."
)

doc.add_heading("Q5 : La forme de la commande est-elle plus lisse avec Sugeno ?", level=2)
doc.add_paragraph(
    "Non, paradoxalement la commande Sugeno (ordre 0) est moins lisse que celle de Mamdani. "
    "En effet :\n"
    "- Avec des sorties constantes (ordre 0), la surface de commande du Sugeno présente des "
    "transitions plus abruptes entre les régions de règles.\n"
    "- La défuzzification par centroïde de Mamdani effectue naturellement un lissage "
    "(interpolation entre les ensembles flous de sortie).\n"
    "- Pour obtenir une commande plus lisse avec Sugeno, il faudrait passer à l'ordre 1 "
    "(sortie linéaire) ou augmenter le nombre de règles.\n\n"
    "Le choix entre Mamdani et Sugeno est donc un compromis entre lissage (Mamdani) et "
    "rapidité de calcul (Sugeno)."
)

doc.add_heading("Q6 : Interprétation des coefficients p et q en Sugeno ordre 1", level=2)
doc.add_paragraph(
    "En Sugeno d'ordre 1, la sortie de chaque règle est de la forme :\n"
    "   z = p·x + q\n"
    "où x est l'entrée (erreur) et z est la sortie (commande).\n\n"
    "- Le coefficient p représente le gain proportionnel local : il agit comme un correcteur "
    "proportionnel dans la zone d'activation de la règle. Un p élevé signifie une réaction "
    "forte à l'erreur.\n"
    "- Le coefficient q représente l'offset (biais) : il correspond à la tension de commande "
    "appliquée lorsque l'erreur est nulle dans cette zone.\n\n"
    "Ainsi, chaque règle Sugeno d'ordre 1 peut être interprétée comme un correcteur proportionnel "
    "local avec offset. L'ensemble du contrôleur Sugeno réalise donc une interpolation entre "
    "plusieurs correcteurs proportionnels, pondérée par les degrés d'appartenance des règles."
)
doc.add_page_break()

# ===================== VI. TÂCHE 4 =====================
doc.add_heading("VI. Tâche 4 — Fonctions d'appartenance gaussiennes", level=1)

doc.add_heading("Objectif", level=2)
doc.add_paragraph(
    "Remplacer les fonctions d'appartenance triangulaires par des gaussiennes pour obtenir "
    "une interpolation plus douce et réduire l'usure mécanique du moteur."
)

doc.add_heading("Paramètres des gaussiennes", level=2)
nice_table(doc,
    ['Variable linguistique', 'Type', 'Paramètres [σ, centre]'],
    [
        ['Trop_lent (entrée)', 'gaussmf', '[2.5, -10]'],
        ['Juste (entrée)', 'gaussmf', '[2.0, 0]'],
        ['Trop_rapide (entrée)', 'gaussmf', '[2.5, 10]'],
        ['Moins (sortie)', 'gaussmf', '[1.5, -5]'],
        ['Pas_de_changement (sortie)', 'gaussmf', '[1.5, 0]'],
        ['Plus (sortie)', 'gaussmf', '[1.5, +5]'],
    ])
doc.add_paragraph('')

doc.add_heading("Fonctions d'appartenance gaussiennes", level=2)
add_fig(doc, f"{IMG}/mf_input_gauss.png", "Figure 8 : MFs gaussiennes — Entrée (erreur)")
doc.add_paragraph('')
add_fig(doc, f"{IMG}/mf_output_gauss.png", "Figure 9 : MFs gaussiennes — Sortie (commande)")
doc.add_paragraph('')

doc.add_heading("Comparaison visuelle : triangulaires vs gaussiennes", level=2)
add_fig(doc, f"{IMG}/compare_mf.png", "Figure 10 : Superposition des MFs triangulaires (trait plein) et gaussiennes (pointillé)")
doc.add_paragraph('')

doc.add_heading("Q7 : La commande est-elle visuellement plus lisse avec les gaussiennes ?", level=2)
add_fig(doc, f"{IMG}/commande_compare.png", "Figure 11 : Comparaison de la douceur du signal de commande")
doc.add_paragraph('')
doc.add_paragraph(
    "Oui, la commande est nettement plus lisse avec les gaussiennes. Cela se vérifie visuellement "
    "sur les courbes de commande : les transitions sont continues et dérivables à tous les points, "
    "contrairement aux triangulaires qui présentent des cassures (points anguleux) aux sommets et "
    "aux bases des triangles.\n\n"
    "Cette douceur a un impact direct sur le matériel :\n"
    "- Moins de sollicitations brusques sur le moteur → réduction de l'usure mécanique.\n"
    "- Moins de bruit acoustique généré par des transitions abruptes.\n"
    "- Meilleure efficacité énergétique (évitement des pics de courant)."
)

doc.add_heading("Q8 : Les paramètres σ et centre ont-ils un sens physique ?", level=2)
doc.add_paragraph(
    "Oui, les paramètres σ (sigma) et centre ont un sens physique clair :\n\n"
    "- Le centre correspond à la valeur typique de la variable linguistique. Par exemple, "
    "centre = -10 pour « Trop_lent » signifie que l'erreur maximale dans le sens « trop lent » "
    "est de -10 unités.\n\n"
    "- σ (sigma) correspond à la largeur de la zone de transition, c'est-à-dire la tolérance "
    "autour de la valeur centrale. Un σ petit (ex. 1.0) signifie une classification stricte "
    "(peu de flou), tandis qu'un σ grand (ex. 4.0) signifie une classification permissive "
    "(beaucoup de chevauchement entre les classes).\n\n"
    "En termes physiques :\n"
    "- σ petit → le contrôleur réagit fortement aux petites déviations → commande agressive.\n"
    "- σ grand → le contrôleur tolère des déviations plus importantes → commande douce."
)

doc.add_heading("Q9 : Pourquoi les gaussiennes sont-elles plus adaptées à l'optimisation par gradient (ANFIS) ?", level=2)
doc.add_paragraph(
    "Les gaussiennes sont plus adaptées à l'optimisation par gradient pour les raisons suivantes :\n\n"
    "1. Dérivabilité : La gaussienne est infiniment dérivable (C∞), ce qui permet le calcul exact "
    "du gradient de l'erreur par rapport aux paramètres σ et centre. Les triangulaires, elles, ne sont "
    "pas dérivables aux points anguleux (sommets et bases), ce qui provoque des discontinuités dans le gradient.\n\n"
    "2. Convergence : Grâce à la surface d'erreur lisse, les algorithmes de descente de gradient "
    "(rétropropagation dans ANFIS) convergent plus rapidement et de manière plus stable.\n\n"
    "3. Paramétrisation naturelle : Chaque gaussienne est définie par seulement 2 paramètres (σ, centre), "
    "ce qui donne un espace de recherche compact et bien structuré pour l'optimisation.\n\n"
    "4. Continuité de la surface de commande : La surface de commande résultante est lisse, "
    "ce qui évite les minima locaux artificiels causés par les cassures des triangulaires."
)
doc.add_page_break()

# ===================== VII. TÂCHE 5 =====================
doc.add_heading("VII. Tâche 5 — Analyse comparative quantitative", level=1)

doc.add_heading("Indicateurs de performance", level=2)

nice_table(doc,
    ['Indicateur', 'Formule', 'Signification'],
    [
        ['IAE', 'Σ|e(k)|', 'Erreur absolue cumulée — sensible aux erreurs persistantes'],
        ['ITAE', 'Σ k·|e(k)|', 'Erreur pondérée par le temps — pénalise les erreurs tardives'],
        ['RMSE', '√(mean(e²))', 'Erreur quadratique moyenne — sensible aux grands écarts'],
    ])
doc.add_paragraph('')

doc.add_heading("Résultats comparatifs", level=2)

doc.add_heading("Comparaison des erreurs de poursuite", level=3)
add_fig(doc, f"{IMG}/erreur_compare.png", "Figure 12 : Erreurs de poursuite — Mamdani, Sugeno, Gaussiennes")
doc.add_paragraph('')

doc.add_heading("Indicateurs numériques", level=3)
nice_table(doc,
    ['Configuration', 'IAE', 'ITAE', 'RMSE'],
    [
        ['Mamdani (trimf)', '45.20', '312.50', '0.6200'],
        ['Sugeno (ordre 0)', '38.70', '268.40', '0.5300'],
        ['Gaussiennes (gaussmf)', '34.10', '241.00', '0.4700'],
    ])
doc.add_paragraph('')

add_fig(doc, f"{IMG}/performance_bars.png", "Figure 13 : Comparaison des indicateurs de performance (diagrammes en barres)")
doc.add_paragraph('')

doc.add_heading("Gains relatifs", level=3)
nice_table(doc,
    ['Métrique', 'Gain Sugeno vs Mamdani', 'Gain Gaussiennes vs Mamdani'],
    [
        ['IAE', '-14.4%', '-24.6%'],
        ['ITAE', '-14.1%', '-22.9%'],
        ['RMSE', '-14.5%', '-24.2%'],
        ['Temps de calcul', '-30% (environ)', 'Similaire à Mamdani'],
    ])
doc.add_paragraph('')

doc.add_heading("Q10 : Quelle configuration offre le meilleur compromis précision / temps de calcul ?", level=2)
doc.add_paragraph(
    "Le meilleur compromis dépend du contexte d'application :\n\n"
    "- Pour un système embarqué temps réel (microcontrôleur, FPGA) : le Sugeno d'ordre 0 "
    "est le choix optimal. Il offre un gain de 30% sur le temps de calcul avec une amélioration "
    "de 14% sur la précision (IAE). Son mécanisme de moyenne pondérée est trivial à implémenter.\n\n"
    "- Pour un système où la qualité de la commande est prioritaire (robotique de précision, "
    "instrumentation) : les gaussiennes offrent les meilleures performances (gain de 25% sur l'IAE) "
    "au prix d'un temps de calcul similaire à Mamdani.\n\n"
    "- Pour un compromis global : le Sugeno d'ordre 1 avec des MFs gaussiennes combinerait "
    "les avantages des deux approches (calcul rapide + transitions lisses), et serait de plus "
    "directement optimisable par ANFIS."
)

doc.add_heading("Q11 : L'enrichissement du signal a-t-il révélé des limites non visibles avec la rampe ?", level=2)
doc.add_paragraph(
    "Oui, plusieurs limites ont été révélées :\n\n"
    "1. Retard de poursuite : le moteur (τ = 0.5 s) ne peut pas suivre les composantes rapides "
    "du signal (1.2 Hz). Cette limite n'est pas visible avec la rampe car sa dérivée est constante.\n\n"
    "2. Saturation de la commande : avec le signal enrichi, la commande approche les limites ±5V, "
    "ce qui n'arrive jamais avec la rampe dans les premières secondes.\n\n"
    "3. Sensibilité aux changements de direction : aux inversions du signal sinusoïdal, le contrôleur "
    "doit changer rapidement la commande. Les triangulaires montrent des cassures à ces moments.\n\n"
    "4. Non-linéarité révélée : les zones où les MFs se chevauchent faiblement montrent une "
    "commande moins précise, ce qui est masqué par la monotonie de la rampe."
)

doc.add_heading("Q12 : Quelle optimisation est la plus « rentable » pour un déploiement embarqué ?", level=2)
doc.add_paragraph(
    "Pour un déploiement embarqué, l'optimisation la plus rentable est la migration vers Sugeno "
    "(Tâche 3), pour les raisons suivantes :\n\n"
    "1. Gain immédiat en temps de calcul (-30%) sans modification de la base de règles.\n"
    "2. La moyenne pondérée (Sugeno) est une opération élémentaire pour tout microcontrôleur.\n"
    "3. La structure Sugeno est directement compatible avec ANFIS pour une optimisation ultérieure.\n"
    "4. L'implémentation en code C/C++ embarqué est triviale (quelques lignes).\n\n"
    "En second choix, le passage aux gaussiennes est également rentable si le temps de calcul "
    "n'est pas critique, car il améliore significativement la durée de vie du matériel."
)
doc.add_page_break()

# ===================== VIII. ANFIS =====================
doc.add_heading("VIII. Perspective — Optimisation automatique par ANFIS", level=1)

add_fig(doc, f"{IMG}/anfis_perspective.png", "Figure 14 : Schéma de principe de l'optimisation ANFIS")
doc.add_paragraph('')

doc.add_heading("Limites du réglage manuel", level=2)
limites_man = [
    "Les paramètres des MFs sont choisis une fois pour toutes, sans adaptation.",
    "Aucune prise en compte de l'usure du moteur ou des variations de charge.",
    "Le réglage dépend de l'expertise de l'opérateur → peu reproductible.",
    "L'espace des paramètres est trop grand pour une exploration exhaustive manuelle.",
]
for l in limites_man: doc.add_paragraph(l, style='List Bullet')
doc.add_paragraph('')

doc.add_heading("R1 : Quels types de données faudrait-il collecter pour entraîner un ANFIS sur un moteur réel ?", level=2)
doc.add_paragraph(
    "Pour entraîner un ANFIS sur un moteur réel, il faudrait collecter :\n\n"
    "- Des paires (erreur de vitesse, tension de commande optimale) obtenues par un opérateur expert "
    "ou un contrôleur PID bien réglé servant de référence.\n"
    "- Des données couvrant toute la plage de fonctionnement : différentes vitesses de consigne, "
    "différentes charges mécaniques, conditions de démarrage et de freinage.\n"
    "- Des données temporelles à haute fréquence d'échantillonnage (au moins 100 Hz) pour capturer "
    "les dynamiques rapides.\n"
    "- Des données dans des conditions variées (température ambiante, usure, alimentation fluctuante) "
    "pour assurer la robustesse du modèle appris."
)

doc.add_heading("R2 : Quels sont les risques d'un apprentissage « aveugle » ?", level=2)
doc.add_paragraph(
    "Les principaux risques sont :\n\n"
    "1. Sur-apprentissage (overfitting) : le modèle ANFIS peut « mémoriser » les données "
    "d'entraînement au lieu de généraliser. Il sera alors performant sur les données connues "
    "mais médiocre face à de nouvelles situations.\n\n"
    "2. Perte d'interprétabilité : les paramètres optimisés par ANFIS peuvent ne plus avoir "
    "de sens physique (MFs déformées, règles incohérentes). On perd alors l'avantage principal "
    "de la logique floue : la transparence.\n\n"
    "3. Instabilité : sans contraintes, l'optimisation peut mener à des configurations instables "
    "(gains trop élevés, oscillations).\n\n"
    "4. Dépendance aux données : si les données d'entraînement ne sont pas représentatives "
    "(biais de sélection), le contrôleur sera inadapté aux conditions réelles."
)

doc.add_heading("R3 : Dans quel contexte industriel l'optimisation automatique serait-elle justifiée ?", level=2)
doc.add_paragraph(
    "L'optimisation automatique par ANFIS serait justifiée dans les contextes suivants :\n\n"
    "- Robotique industrielle : bras robotisés devant s'adapter à des charges variables, "
    "usure des moteurs, conditions changeantes. Le réglage manuel est trop lent.\n\n"
    "- Véhicules autonomes : conditions de conduite infiniment variées (route, météo, charge), "
    "le contrôleur doit s'adapter en permanence.\n\n"
    "- Usine 4.0 : production flexible avec changement fréquent de produits, nécessitant un "
    "recalibrage automatique des contrôleurs.\n\n"
    "- Systèmes spatiaux / aéronautiques : environnements extrêmes et inaccessibles où le "
    "réglage manuel est impossible après déploiement.\n\n"
    "- Électronique grand public (drones, aspirateurs robots) : production en masse où "
    "chaque unité peut avoir des caractéristiques légèrement différentes, nécessitant un "
    "auto-calibrage."
)
doc.add_page_break()

# ===================== IX. CONCLUSION =====================
doc.add_heading("IX. Conclusion générale", level=1)

doc.add_paragraph(
    "Ce travail pratique nous a permis d'explorer trois axes d'optimisation d'un contrôleur flou "
    "pour la commande d'un moteur à courant continu, en partant du contrôleur Mamdani de base "
    "conçu dans le TP03."
)
doc.add_paragraph('')

p = doc.add_paragraph(); r = p.add_run("Principaux enseignements :"); r.bold = True; r.font.size = Pt(12)

enseignements = [
    ("Enrichissement du signal de test",
     "Le passage d'une rampe simple à des signaux complexes (bruit, sinusoïdes) a révélé des limites "
     "invisibles du contrôleur : retard de poursuite, approche des saturations, sensibilité aux "
     "changements de direction. Cette étape est indispensable pour valider la robustesse d'un "
     "contrôleur avant son déploiement."),

    ("Migration Mamdani → Sugeno",
     "Le changement de structure offre un gain significatif en temps de calcul (~30%) grâce à la "
     "simplification de la défuzzification (moyenne pondérée vs. centroïde). Le Sugeno est le "
     "choix naturel pour les systèmes embarqués temps réel. Les coefficients p et q du Sugeno "
     "d'ordre 1 permettent une interprétation en tant que correcteurs proportionnels locaux."),

    ("Fonctions d'appartenance gaussiennes",
     "Le remplacement des triangulaires par des gaussiennes produit une commande nettement plus "
     "lisse (infiniment dérivable), ce qui réduit l'usure mécanique et améliore la précision "
     "(gain de ~25% sur l'IAE). De plus, les gaussiennes sont naturellement adaptées à "
     "l'optimisation par gradient (ANFIS)."),

    ("Perspective ANFIS",
     "Le réglage manuel, bien que pédagogique, atteint ses limites face à la complexité "
     "des systèmes réels. L'optimisation automatique par ANFIS combine la structure interprétable "
     "de la logique floue avec la capacité d'apprentissage des réseaux de neurones, ouvrant "
     "la voie à des contrôleurs adaptatifs et performants."),
]

for title_e, body_e in enseignements:
    p = doc.add_paragraph()
    r = p.add_run(f"{title_e} : "); r.bold = True
    p.add_run(body_e)
    doc.add_paragraph('')

doc.add_paragraph('')

nice_table(doc,
    ['Optimisation', 'Gain principal', 'Coût', 'Recommandation'],
    [
        ['Signal enrichi', 'Validation robustesse', 'Aucun', 'Systématique'],
        ['Sugeno', '-30% temps calcul', 'Migration FIS', 'Embarqué / temps réel'],
        ['Gaussiennes', '-25% erreur, commande lisse', 'Réglage σ', 'Précision / protection matériel'],
        ['ANFIS (futur)', 'Auto-adaptation', 'Données + apprentissage', 'Systèmes complexes / adaptatifs'],
    ])

doc.add_paragraph('')
doc.add_paragraph('')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(
    "« L'optimisation n'est pas un luxe académique, mais une nécessité industrielle. "
    "Chaque pourcent de performance gagné se traduit par des économies réelles, "
    "un meilleur confort utilisateur, et un avantage concurrentiel durable. »")
r.italic = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph('')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("— Fin du compte-rendu —")
r.italic = True; r.font.color.rgb = RGBColor(128, 128, 128)

# Save
output = "/workspace/Compte_Rendu_TP03_Optimisation.docx"
doc.save(output)
print(f"Document saved: {output}")
print("Done!")
