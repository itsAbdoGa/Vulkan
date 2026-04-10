#!/usr/bin/env python3
"""
Script to generate the TP03 Word document with answers and illustrations.
Transient Analysis of Power Transformers - ATP-EMTP
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import math

IMG_DIR = "/workspace/images"
os.makedirs(IMG_DIR, exist_ok=True)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def set_cell_shading(cell, color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)

def add_table_with_style(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(10)
        set_cell_shading(cell, "4472C4")
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(val)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.size = Pt(10)
    return table


# ============================================================
# FIGURE 1: Single-phase transformer equivalent circuit
# ============================================================

def create_fig_task1_circuit():
    fig, ax = plt.subplots(1, 1, figsize=(12, 5))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Schéma équivalent du transformateur monophasé - Essai de court-circuit", fontsize=13, fontweight='bold')

    # Source (LS side)
    ax.annotate('', xy=(1.5, 4), xytext=(0.5, 4), arrowprops=dict(arrowstyle='->', lw=2))
    ax.text(0.2, 4.3, 'Source\n(BT)', fontsize=9, ha='center', va='bottom', color='blue')

    # R1 (LS resistance)
    rect1 = FancyBboxPatch((1.8, 3.7), 1.2, 0.6, boxstyle="round,pad=0.05", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect1)
    ax.text(2.4, 4.0, 'R₁', fontsize=10, ha='center', va='center', fontweight='bold')

    # L1 (LS leakage)
    ax.plot([3.0, 3.3], [4.0, 4.0], 'k-', lw=2)
    for i in range(4):
        theta = np.linspace(0, np.pi, 30)
        x = 3.3 + i * 0.3 + 0.15 + 0.15 * np.cos(theta)
        y = 4.0 + 0.15 * np.sin(theta)
        ax.plot(x, y, 'b-', lw=2)
    ax.plot([4.5, 4.8], [4.0, 4.0], 'k-', lw=2)
    ax.text(3.9, 4.5, 'L₁', fontsize=10, ha='center', va='center', fontweight='bold', color='blue')

    # Ideal transformer
    ax.plot([4.8, 4.8], [3.3, 4.7], 'k-', lw=2)
    ax.plot([5.0, 5.0], [3.3, 4.7], 'k-', lw=2)
    for i in range(5):
        theta = np.linspace(-np.pi/2, np.pi/2, 30)
        y1 = 3.4 + i * 0.26 + 0.13
        x1 = 4.8 - 0.13 * np.cos(theta)
        y1_arr = y1 + 0.13 * np.sin(theta)
        ax.plot(x1, y1_arr, 'b-', lw=1.5)
    for i in range(5):
        theta = np.linspace(np.pi/2, 3*np.pi/2, 30)
        y1 = 3.4 + i * 0.26 + 0.13
        x1 = 5.0 - 0.13 * np.cos(theta)
        y1_arr = y1 + 0.13 * np.sin(theta)
        ax.plot(x1, y1_arr, 'r-', lw=1.5)

    # Magnetizing branch (Rm, Lm)
    ax.plot([4.8, 4.8], [2.0, 3.3], 'k-', lw=1)
    ax.plot([4.8, 4.3], [2.6, 2.6], 'k-', lw=1)
    ax.plot([4.8, 5.3], [2.6, 2.6], 'k-', lw=1)
    rect_rm = FancyBboxPatch((4.0, 1.9), 0.6, 0.5, boxstyle="round,pad=0.02", facecolor='#90EE90', edgecolor='black', lw=1)
    ax.add_patch(rect_rm)
    ax.text(4.3, 2.15, 'Rm', fontsize=8, ha='center', va='center')
    rect_lm = FancyBboxPatch((5.0, 1.9), 0.6, 0.5, boxstyle="round,pad=0.02", facecolor='#ADD8E6', edgecolor='black', lw=1)
    ax.add_patch(rect_lm)
    ax.text(5.3, 2.15, 'Lm', fontsize=8, ha='center', va='center')
    ax.plot([4.3, 4.3], [1.9, 1.5], 'k-', lw=1)
    ax.plot([5.3, 5.3], [1.9, 1.5], 'k-', lw=1)
    ax.plot([4.3, 5.3], [1.5, 1.5], 'k-', lw=1)
    ax.text(4.8, 1.2, 'Branche\nmagnétisante', fontsize=7, ha='center', va='center', color='green', style='italic')

    # R2 (HS resistance)
    ax.plot([5.0, 5.5], [4.0, 4.0], 'k-', lw=2)
    rect2 = FancyBboxPatch((5.5, 3.7), 1.2, 0.6, boxstyle="round,pad=0.05", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect2)
    ax.text(6.1, 4.0, 'R₂', fontsize=10, ha='center', va='center', fontweight='bold')

    # L2 (HS leakage)
    ax.plot([6.7, 7.0], [4.0, 4.0], 'k-', lw=2)
    for i in range(4):
        theta = np.linspace(0, np.pi, 30)
        x = 7.0 + i * 0.3 + 0.15 + 0.15 * np.cos(theta)
        y = 4.0 + 0.15 * np.sin(theta)
        ax.plot(x, y, 'r-', lw=2)
    ax.plot([8.2, 8.5], [4.0, 4.0], 'k-', lw=2)
    ax.text(7.6, 4.5, 'L₂', fontsize=10, ha='center', va='center', fontweight='bold', color='red')

    # Zo (Load)
    rect_zo = FancyBboxPatch((9.0, 3.5), 1.2, 1.0, boxstyle="round,pad=0.05", facecolor='#DDA0DD', edgecolor='black', lw=1.5)
    ax.add_patch(rect_zo)
    ax.text(9.6, 4.0, 'Zo', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.plot([8.5, 9.0], [4.0, 4.0], 'k-', lw=2)

    # RF (Short-circuit resistance)
    ax.plot([10.2, 11.0], [4.0, 4.0], 'k-', lw=2)
    rect_rf = FancyBboxPatch((11.0, 3.5), 1.2, 1.0, boxstyle="round,pad=0.05", facecolor='#FF6347', edgecolor='black', lw=1.5)
    ax.add_patch(rect_rf)
    ax.text(11.6, 4.0, 'RF\n0.5 Ω', fontsize=9, ha='center', va='center', fontweight='bold', color='white')
    ax.text(11.6, 3.0, 'Court-circuit\n(côté HT)', fontsize=8, ha='center', va='center', color='red', style='italic')

    # Switch
    ax.plot([12.2, 12.8], [4.0, 4.3], 'k-', lw=2)
    ax.plot(12.2, 4.0, 'ko', markersize=5)
    ax.plot([12.8, 13.0], [4.0, 4.0], 'k-', lw=2)
    ax.text(12.5, 4.6, 'S', fontsize=10, ha='center', va='center', fontweight='bold')

    # Return path
    ax.plot([0.5, 0.5], [2.0, 4.0], 'k-', lw=2)
    ax.plot([0.5, 13.0], [2.0, 2.0], 'k-', lw=2)
    ax.plot([13.0, 13.0], [2.0, 4.0], 'k-', lw=2)

    # Labels
    ax.text(0.5, 5.0, 'Côté BT (15.75 kV)', fontsize=10, ha='center', color='blue', fontweight='bold')
    ax.text(11.0, 5.0, 'Côté HT (400 kV)', fontsize=10, ha='center', color='red', fontweight='bold')

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task1_circuit.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 2: Transformer parameter calculation diagram
# ============================================================

def create_fig_task1_params():
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    # Short-circuit test equivalent
    ax = axes[0]
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Essai de court-circuit\n(Zcc vue du primaire)", fontsize=11, fontweight='bold')

    ax.plot([1, 2], [4, 4], 'b-', lw=2)
    rect_r = FancyBboxPatch((2, 3.6), 1.5, 0.8, boxstyle="round,pad=0.05", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_r)
    ax.text(2.75, 4.0, 'Rcc', fontsize=11, ha='center', va='center', fontweight='bold')
    ax.plot([3.5, 4.5], [4, 4], 'b-', lw=2)
    rect_x = FancyBboxPatch((4.5, 3.6), 1.5, 0.8, boxstyle="round,pad=0.05", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_x)
    ax.text(5.25, 4.0, 'Xcc', fontsize=11, ha='center', va='center', fontweight='bold')
    ax.plot([6, 7], [4, 4], 'b-', lw=2)
    ax.plot([7, 7], [3, 4], 'b-', lw=2)
    ax.plot([1, 1], [3, 4], 'b-', lw=2)
    ax.plot([1, 7], [3, 3], 'b-', lw=2)

    ax.text(4, 2.0, r'$Z_{cc} = R_{cc} + jX_{cc}$', fontsize=12, ha='center')
    ax.text(4, 1.3, r'$R_{cc} = \frac{P_{cc}}{I_n^2}$    $X_{cc} = \sqrt{Z_{cc}^2 - R_{cc}^2}$', fontsize=10, ha='center')

    # Open-circuit test equivalent
    ax = axes[1]
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Essai à vide\n(Branche magnétisante)", fontsize=11, fontweight='bold')

    ax.plot([2, 4], [4, 4], 'r-', lw=2)
    ax.plot([4, 4], [4, 2], 'r-', lw=2)
    ax.plot([4, 3.5], [2, 2], 'r-', lw=1.5)
    ax.plot([4, 4.5], [2, 2], 'r-', lw=1.5)

    rect_rm = FancyBboxPatch((3.0, 1.0), 1.0, 0.8, boxstyle="round,pad=0.05", facecolor='#90EE90', edgecolor='black', lw=1.5)
    ax.add_patch(rect_rm)
    ax.text(3.5, 1.4, 'Rm', fontsize=10, ha='center', va='center', fontweight='bold')

    rect_lm = FancyBboxPatch((4.0, 1.0), 1.0, 0.8, boxstyle="round,pad=0.05", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_lm)
    ax.text(4.5, 1.4, 'Xm', fontsize=10, ha='center', va='center', fontweight='bold')

    ax.plot([3.5, 3.5], [1.0, 0.5], 'r-', lw=1.5)
    ax.plot([4.5, 4.5], [1.0, 0.5], 'r-', lw=1.5)
    ax.plot([3.5, 4.5], [0.5, 0.5], 'r-', lw=1.5)
    ax.plot([2, 2], [0.5, 4], 'r-', lw=1.5)
    ax.plot([2, 3.5], [0.5, 0.5], 'r-', lw=1.5)

    ax.text(4, 5.0, r'$R_m = \frac{U_n^2}{P_{fer}}$    $X_m = \frac{U_n}{I_0}$', fontsize=10, ha='center')

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task1_params.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 3: CT model circuit
# ============================================================

def create_fig_task2_ct():
    fig, ax = plt.subplots(1, 1, figsize=(12, 5))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle du transformateur de courant (TC) 5P20 dans ATPDraw", fontsize=13, fontweight='bold')

    # Source impedance Zs
    ax.text(0.5, 4.3, 'Source\nI₁ = 10 kA', fontsize=8, ha='center', color='blue')
    ax.annotate('', xy=(1.5, 4), xytext=(0.5, 4), arrowprops=dict(arrowstyle='->', lw=2, color='blue'))
    rect_zs = FancyBboxPatch((1.5, 3.6), 1.2, 0.8, boxstyle="round,pad=0.05", facecolor='#B0C4DE', edgecolor='black', lw=1.5)
    ax.add_patch(rect_zs)
    ax.text(2.1, 4.0, 'Zs', fontsize=10, ha='center', va='center', fontweight='bold')

    # CT primary (N1 turns)
    ax.plot([2.7, 4.0], [4.0, 4.0], 'k-', lw=2)
    ax.plot([4.0, 4.0], [3.0, 5.0], 'k-', lw=2)
    ax.plot([4.3, 4.3], [3.0, 5.0], 'k-', lw=2)
    ax.text(4.15, 5.3, 'TC\n500:1', fontsize=9, ha='center', fontweight='bold', color='purple')

    for i in range(4):
        theta = np.linspace(-np.pi/2, np.pi/2, 30)
        y = 3.2 + i * 0.45 + 0.2
        x = 4.0 - 0.15 * np.cos(theta)
        y_arr = y + 0.2 * np.sin(theta)
        ax.plot(x, y_arr, 'b-', lw=1.5)
    for i in range(4):
        theta = np.linspace(np.pi/2, 3*np.pi/2, 30)
        y = 3.2 + i * 0.45 + 0.2
        x = 4.3 - 0.15 * np.cos(theta)
        y_arr = y + 0.2 * np.sin(theta)
        ax.plot(x, y_arr, 'r-', lw=1.5)

    # Secondary R2
    ax.plot([4.3, 5.0], [4.0, 4.0], 'k-', lw=2)
    rect_r2 = FancyBboxPatch((5.0, 3.6), 1.0, 0.8, boxstyle="round,pad=0.05", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_r2)
    ax.text(5.5, 4.0, 'R₂\n4.5Ω', fontsize=9, ha='center', va='center', fontweight='bold')

    # Magnetizing branch (Element 98)
    ax.plot([6.0, 7.0], [4.0, 4.0], 'k-', lw=2)
    ax.plot([7.0, 7.0], [4.0, 2.5], 'k--', lw=1.5, color='green')

    rect_98 = FancyBboxPatch((6.5, 1.5), 1.0, 1.0, boxstyle="round,pad=0.05", facecolor='#98FB98', edgecolor='green', lw=2)
    ax.add_patch(rect_98)
    ax.text(7.0, 2.0, 'Type\n98', fontsize=9, ha='center', va='center', fontweight='bold', color='green')

    # DC source (type 11) for initial flux
    rect_dc = FancyBboxPatch((6.5, 0.5), 1.0, 0.7, boxstyle="round,pad=0.05", facecolor='#FFA07A', edgecolor='black', lw=1)
    ax.add_patch(rect_dc)
    ax.text(7.0, 0.85, 'DC\n(Type 11)', fontsize=7, ha='center', va='center', fontweight='bold')
    ax.plot([7.0, 7.0], [1.2, 1.5], 'k-', lw=1)

    ax.text(8.0, 1.5, 'Branche\nmagnétisante\n(élément 98 +\nsource DC)', fontsize=7, ha='left', color='green', style='italic')

    # Load Zobc
    ax.plot([7.0, 7.0], [4.0, 4.0], 'k-', lw=0)
    ax.plot([7.0, 9.0], [4.0, 4.0], 'k-', lw=2)
    rect_zobc = FancyBboxPatch((9.0, 3.3), 1.5, 1.4, boxstyle="round,pad=0.05", facecolor='#DDA0DD', edgecolor='black', lw=1.5)
    ax.add_patch(rect_zobc)
    ax.text(9.75, 4.0, 'Zobc\n10+j17.3Ω', fontsize=9, ha='center', va='center', fontweight='bold')

    # Return
    ax.plot([0.5, 0.5], [2.0, 4.0], 'k-', lw=2)
    ax.plot([0.5, 10.5], [2.0, 2.0], 'k-', lw=2)
    ax.plot([10.5, 10.5], [2.0, 4.0], 'k-', lw=2)

    ax.text(6.0, 5.5, 'Éléments du TC (en pointillé vert)', fontsize=9, ha='center', color='green', style='italic')

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task2_ct_model.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 4: CT Magnetizing characteristic B-H / lambda-i
# ============================================================

def create_fig_task2_bh():
    # B-H data from the PDF
    i_data = [0.0143, 0.0382, 0.0573, 0.0955, 0.1909, 0.7637, 3.8184, 28.6378]
    lambda_data = [0.1440, 1.4400, 2.0160, 2.3184, 2.4912, 2.6928, 2.8368, 2.9664]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

    # Lambda-i curve
    ax1.plot(i_data, lambda_data, 'b-o', lw=2, markersize=6, label=r'$\lambda = f(i_\mu)$')
    ax1.set_xlabel(r'Courant magnétisant $i_\mu$ (A)', fontsize=11)
    ax1.set_ylabel(r'Flux $\lambda$ (Vs)', fontsize=11)
    ax1.set_title(r'Caractéristique magnétisante $\lambda = f(i_\mu)$', fontsize=12, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=10)
    ax1.set_xscale('log')

    # B-H curve reconstruction
    N2 = 500
    S = 2.88e-3
    l_core = 0.675
    B_data = [lam / (N2 * S) for lam in lambda_data]
    H_data = [i * N2 / l_core for i in i_data]

    ax2.plot(H_data, B_data, 'r-s', lw=2, markersize=6, label='B = f(H)')
    ax2.set_xlabel('H (A/m)', fontsize=11)
    ax2.set_ylabel('B (T)', fontsize=11)
    ax2.set_title('Courbe B-H du noyau du TC', fontsize=12, fontweight='bold')
    ax2.grid(True, alpha=0.3)
    ax2.legend(fontsize=10)
    ax2.set_xscale('log')

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task2_bh_curve.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 5: Three-phase transformer model (Task 3)
# ============================================================

def create_fig_task3_circuit():
    fig, ax = plt.subplots(1, 1, figsize=(14, 7))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 8)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle STC du transformateur triphasé 24kV/230kV (Δ-Yn) - Court-circuit au secondaire",
                  fontsize=12, fontweight='bold')

    # Three-phase source (230 kV)
    for phase_idx, (color, name, y_pos) in enumerate([('red', 'A', 6.5), ('blue', 'B', 4.5), ('green', 'C', 2.5)]):
        # Source
        circle = plt.Circle((1.5, y_pos), 0.4, fill=False, edgecolor=color, lw=2)
        ax.add_patch(circle)
        ax.text(1.5, y_pos, f'~', fontsize=14, ha='center', va='center', color=color)
        ax.text(0.5, y_pos, f'Phase {name}\n230kV', fontsize=7, ha='center', va='center', color=color)

        # Line to breaker
        ax.plot([1.9, 3.0], [y_pos, y_pos], color=color, lw=2)

        # Circuit breaker S1
        ax.plot([3.0, 3.5], [y_pos, y_pos + 0.3], color=color, lw=2)
        ax.plot(3.0, y_pos, 'o', color=color, markersize=4)
        ax.plot(3.5, y_pos + 0.3, 'o', color=color, markersize=4)

        # Line to transformer
        ax.plot([3.5, 5.0], [y_pos + 0.3, y_pos + 0.3], color=color, lw=1)
        ax.plot([5.0, 5.0], [y_pos + 0.3, y_pos], color=color, lw=1)
        ax.plot([5.0, 6.0], [y_pos, y_pos], color=color, lw=2)

    # Transformer box
    rect_tx = FancyBboxPatch((6.0, 1.5), 3.0, 6.0, boxstyle="round,pad=0.1", facecolor='#E8E8E8', edgecolor='black', lw=2)
    ax.add_patch(rect_tx)
    ax.text(7.5, 7.0, 'Transformateur STC\n24kV:230kV\n150 MVA\nΔ / Y-gnd', fontsize=9, ha='center', va='center', fontweight='bold')

    # Delta symbol on primary side
    ax.text(6.5, 3.5, 'Δ', fontsize=20, ha='center', va='center', color='blue')

    # Y symbol on secondary side
    ax.text(8.5, 3.5, 'Y', fontsize=20, ha='center', va='center', color='red')

    # Secondary lines
    for phase_idx, (color, name, y_pos) in enumerate([('red', 'a', 6.5), ('blue', 'b', 4.5), ('green', 'c', 2.5)]):
        ax.plot([9.0, 10.5], [y_pos, y_pos], color=color, lw=2)

        # Short-circuit switch
        ax.plot([10.5, 11.0], [y_pos, y_pos + 0.2], color='red', lw=2)
        ax.plot(10.5, y_pos, 'o', color='red', markersize=4)

        # Short-circuit to ground
        ax.plot([11.0, 12.0], [y_pos, y_pos], 'r--', lw=1.5)
        ax.plot([12.0, 12.0], [y_pos, y_pos - 0.3], 'k-', lw=2)
        ax.plot([11.7, 12.3], [y_pos - 0.3, y_pos - 0.3], 'k-', lw=2)
        ax.plot([11.8, 12.2], [y_pos - 0.5, y_pos - 0.5], 'k-', lw=1.5)
        ax.plot([11.9, 12.1], [y_pos - 0.7, y_pos - 0.7], 'k-', lw=1)

    ax.text(11.0, 7.5, 'Court-circuit triphasé\nt = 0.1 s', fontsize=10, ha='center', color='red', fontweight='bold')
    ax.text(3.3, 7.5, 'Disjoncteur S1', fontsize=9, ha='center', color='black')

    # Ground connection (Y-gnd)
    ax.plot([8.5, 8.5], [1.5, 1.0], 'k-', lw=2)
    ax.plot([8.2, 8.8], [1.0, 1.0], 'k-', lw=2)
    ax.plot([8.3, 8.7], [0.8, 0.8], 'k-', lw=1.5)
    ax.plot([8.4, 8.6], [0.6, 0.6], 'k-', lw=1)

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task3_circuit.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 6: Expected short-circuit current waveform (Task 3)
# ============================================================

def create_fig_task3_waveform():
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))

    t = np.linspace(0, 0.3, 3000)
    f = 60
    omega = 2 * np.pi * f

    # Before short-circuit (0 to 0.1s): normal current
    # After short-circuit (0.1 to 0.3s): large transient current

    I_normal = 376  # nominal current at 230kV side (A)
    I_sc_peak = 376 * 15  # short-circuit peak current
    tau = 0.05  # time constant for DC offset decay

    i_a = np.zeros_like(t)
    i_b = np.zeros_like(t)
    i_c = np.zeros_like(t)

    for idx, ti in enumerate(t):
        if ti < 0.1:
            i_a[idx] = I_normal * np.sqrt(2) * np.sin(omega * ti)
            i_b[idx] = I_normal * np.sqrt(2) * np.sin(omega * ti - 2*np.pi/3)
            i_c[idx] = I_normal * np.sqrt(2) * np.sin(omega * ti + 2*np.pi/3)
        else:
            dt = ti - 0.1
            dc_offset = I_sc_peak * np.exp(-dt / tau)
            i_a[idx] = I_sc_peak * np.sin(omega * ti) + dc_offset * 0.8
            i_b[idx] = I_sc_peak * np.sin(omega * ti - 2*np.pi/3) + dc_offset * 0.3
            i_c[idx] = I_sc_peak * np.sin(omega * ti + 2*np.pi/3) - dc_offset * 0.5

    ax1.plot(t * 1000, i_a / 1000, 'r-', lw=1, label='Phase A')
    ax1.plot(t * 1000, i_b / 1000, 'b-', lw=1, label='Phase B')
    ax1.plot(t * 1000, i_c / 1000, 'g-', lw=1, label='Phase C')
    ax1.axvline(x=100, color='k', linestyle='--', lw=1.5, label='Court-circuit (t=0.1s)')
    ax1.set_xlabel('Temps (ms)', fontsize=11)
    ax1.set_ylabel('Courant (kA)', fontsize=11)
    ax1.set_title('Courant dans le disjoncteur S1 - Court-circuit triphasé', fontsize=12, fontweight='bold')
    ax1.legend(fontsize=9)
    ax1.grid(True, alpha=0.3)

    # RMS calculation
    window = int(1 / (f * (t[1] - t[0])))  # one cycle window
    i_rms_a = np.zeros_like(t)
    for idx in range(window, len(t)):
        i_rms_a[idx] = np.sqrt(np.mean(i_a[idx-window:idx]**2))

    ax2.plot(t * 1000, i_rms_a / 1000, 'r-', lw=2, label='RMS Phase A')
    ax2.axvline(x=100, color='k', linestyle='--', lw=1.5, label='Court-circuit')
    ax2.set_xlabel('Temps (ms)', fontsize=11)
    ax2.set_ylabel('Courant RMS (kA)', fontsize=11)
    ax2.set_title('Valeur RMS du courant dans le disjoncteur S1', fontsize=12, fontweight='bold')
    ax2.legend(fontsize=9)
    ax2.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task3_waveform.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 7: Auto-transformer energization (Task 4)
# ============================================================

def create_fig_task4_circuit():
    fig, ax = plt.subplots(1, 1, figsize=(14, 7))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 8)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Schéma d'énergisation de l'autotransformateur 400/132/18 kV (Yyd)",
                  fontsize=12, fontweight='bold')

    # Source 420kV
    for idx, (color, name, y_pos) in enumerate([('red', 'A', 6.5), ('blue', 'B', 4.5), ('green', 'C', 2.5)]):
        circle = plt.Circle((1.5, y_pos), 0.4, fill=False, edgecolor=color, lw=2)
        ax.add_patch(circle)
        ax.text(1.5, y_pos, '~', fontsize=14, ha='center', va='center', color=color)
        ax.text(0.5, y_pos, f'{name}', fontsize=9, ha='center', va='center', color=color, fontweight='bold')

        # Switch
        ax.plot([1.9, 3.0], [y_pos, y_pos], color=color, lw=2)
        ax.plot([3.0, 3.6], [y_pos, y_pos + 0.2], color=color, lw=2)
        ax.plot(3.0, y_pos, 'o', color=color, markersize=4)

        # Source impedance
        ax.plot([3.6, 4.5], [y_pos, y_pos], color=color, lw=2)
        rect = FancyBboxPatch((4.5, y_pos - 0.25), 1.0, 0.5, boxstyle="round,pad=0.03", facecolor='#B0C4DE', edgecolor='black', lw=1)
        ax.add_patch(rect)
        ax.text(5.0, y_pos, 'Zs', fontsize=8, ha='center', va='center')
        ax.plot([5.5, 6.5], [y_pos, y_pos], color=color, lw=2)

    ax.text(1.5, 7.5, 'Source\n420 kV, 60 Hz', fontsize=9, ha='center', fontweight='bold')

    # Transformer box
    rect_tx = FancyBboxPatch((6.5, 1.5), 3.5, 6.0, boxstyle="round,pad=0.1",
                              facecolor='#FFF8DC', edgecolor='black', lw=2)
    ax.add_patch(rect_tx)

    ax.text(8.25, 6.8, 'Autotransformateur', fontsize=10, ha='center', fontweight='bold')
    ax.text(8.25, 6.2, '400/132/18 kV', fontsize=9, ha='center')
    ax.text(8.25, 5.7, 'Couplage Yyd', fontsize=9, ha='center')
    ax.text(8.25, 5.0, 'Type-96\n(inducteurs\nhystérétiques)', fontsize=8, ha='center', color='purple', style='italic')

    ax.text(7.0, 3.5, 'Y', fontsize=18, ha='center', va='center', color='blue')
    ax.text(8.25, 3.5, 'y', fontsize=18, ha='center', va='center', color='red')
    ax.text(9.5, 3.5, 'd', fontsize=18, ha='center', va='center', color='green')

    # 400kV label
    ax.text(6.5, 1.0, '400 kV', fontsize=9, ha='center', color='blue', fontweight='bold')
    # 132kV label
    ax.text(8.25, 1.0, '132 kV', fontsize=9, ha='center', color='red', fontweight='bold')
    # 18kV label
    ax.text(10.0, 1.0, '18 kV', fontsize=9, ha='center', color='green', fontweight='bold')

    # Secondary 132kV (open)
    for idx, (color, y_pos) in enumerate([('red', 6.5), ('red', 4.5), ('red', 2.5)]):
        ax.plot([10.0, 11.5], [y_pos, y_pos], 'r-', lw=2)
        ax.plot([11.5, 12.0], [y_pos - 0.1, y_pos + 0.1], 'r-', lw=3)
        ax.text(11.8, y_pos + 0.3, 'ouvert', fontsize=7, ha='center', color='red')

    # Tertiary 18kV (open)
    for idx, (color, y_pos) in enumerate([('green', 6.0), ('green', 4.0), ('green', 2.0)]):
        ax.plot([10.0, 11.5], [y_pos, y_pos], 'g--', lw=1.5)
        ax.plot([11.5, 12.0], [y_pos - 0.1, y_pos + 0.1], 'g-', lw=3)
        ax.text(11.8, y_pos + 0.3, 'ouvert', fontsize=7, ha='center', color='green')

    ax.text(13.0, 5.5, '132 kV\n(à vide)', fontsize=9, ha='center', color='red', fontweight='bold')
    ax.text(13.0, 3.0, '18 kV\n(à vide)', fontsize=9, ha='center', color='green', fontweight='bold')

    # Timing info
    ax.text(3.3, 1.0, 'Désenergisation: t = 45 ms\nRéenergisation:\n  Phase A: t = 73.5 ms\n  Phases B,C: t = 78.5 ms',
            fontsize=8, ha='left', color='purple', fontweight='bold',
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task4_circuit.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 8: Inrush current waveform (Task 4)
# ============================================================

def create_fig_task4_inrush():
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))

    t = np.linspace(0, 0.15, 3000)
    f = 60
    omega = 2 * np.pi * f
    t_deenerg = 0.045
    t_reenerg_a = 0.0735
    t_reenerg_bc = 0.0785

    i_mag_a = np.zeros_like(t)
    i_mag_b = np.zeros_like(t)
    i_mag_c = np.zeros_like(t)

    I_mag_nom = 5.0  # nominal magnetizing current (A)
    I_inrush_peak = 200.0  # peak inrush current (A)
    tau_inrush = 0.03

    for idx, ti in enumerate(t):
        if ti < t_deenerg:
            # Steady state - small magnetizing current
            i_mag_a[idx] = I_mag_nom * np.sin(omega * ti)
            i_mag_b[idx] = I_mag_nom * np.sin(omega * ti - 2*np.pi/3)
            i_mag_c[idx] = I_mag_nom * np.sin(omega * ti + 2*np.pi/3)
        elif ti < t_reenerg_a:
            # De-energized: currents decay quickly to zero
            dt = ti - t_deenerg
            i_mag_a[idx] = I_mag_nom * np.sin(omega * t_deenerg) * np.exp(-dt / 0.002)
            i_mag_b[idx] = I_mag_nom * np.sin(omega * t_deenerg - 2*np.pi/3) * np.exp(-dt / 0.002)
            i_mag_c[idx] = I_mag_nom * np.sin(omega * t_deenerg + 2*np.pi/3) * np.exp(-dt / 0.002)
        else:
            # Re-energization with inrush
            if ti >= t_reenerg_a:
                dt_a = ti - t_reenerg_a
                dc_a = I_inrush_peak * np.exp(-dt_a / tau_inrush)
                i_mag_a[idx] = I_inrush_peak * 0.3 * np.sin(omega * ti) + dc_a * np.maximum(np.sin(omega * ti), 0)
            if ti >= t_reenerg_bc:
                dt_bc = ti - t_reenerg_bc
                dc_bc = I_inrush_peak * 0.7 * np.exp(-dt_bc / tau_inrush)
                i_mag_b[idx] = I_inrush_peak * 0.2 * np.sin(omega * ti - 2*np.pi/3) + dc_bc * 0.5 * np.maximum(np.sin(omega * ti - 2*np.pi/3), 0)
                i_mag_c[idx] = I_inrush_peak * 0.2 * np.sin(omega * ti + 2*np.pi/3) - dc_bc * 0.3 * np.maximum(-np.sin(omega * ti + 2*np.pi/3), 0)

    ax1.plot(t * 1000, i_mag_a, 'r-', lw=1, label='Phase A')
    ax1.plot(t * 1000, i_mag_b, 'b-', lw=1, label='Phase B')
    ax1.plot(t * 1000, i_mag_c, 'g-', lw=1, label='Phase C')
    ax1.axvline(x=45, color='orange', linestyle='--', lw=1.5, label='Désenergisation (45 ms)')
    ax1.axvline(x=73.5, color='purple', linestyle='--', lw=1.5, label='Réenergisation A (73.5 ms)')
    ax1.axvline(x=78.5, color='brown', linestyle=':', lw=1.5, label='Réenergisation B,C (78.5 ms)')
    ax1.set_xlabel('Temps (ms)', fontsize=11)
    ax1.set_ylabel('Courant magnétisant (A)', fontsize=11)
    ax1.set_title("Courant magnétisant et courant d'appel (inrush) lors de la réenergisation", fontsize=12, fontweight='bold')
    ax1.legend(fontsize=8, loc='upper right')
    ax1.grid(True, alpha=0.3)

    # Zoom on steady-state
    t_ss = np.linspace(0, 0.04, 500)
    i_ss = I_mag_nom * np.sin(omega * t_ss)
    ax2.plot(t_ss * 1000, i_ss, 'r-', lw=2, label='Courant magnétisant (régime permanent)')
    ax2.set_xlabel('Temps (ms)', fontsize=11)
    ax2.set_ylabel('Courant (A)', fontsize=11)
    ax2.set_title('Courant magnétisant en régime permanent (avant désenergisation)', fontsize=12, fontweight='bold')
    ax2.legend(fontsize=9)
    ax2.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task4_inrush.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 9: Hysteresis curve (Task 4)
# ============================================================

def create_fig_task4_hysteresis():
    # Data from the PDF
    current = [-36.825, -24.55, -11.0475, -4.91, -1.84125, 0.61375, 2.148125,
               3.55975, 4.29625, 4.91, 6.1375, 6.75125, 8.5925, 11.0475,
               13.37975, 17.491875, 23.93625, 32.835625, 42.9625, 61.375, 98.2, 135.025]
    flux = [-94.9129412, -94.3411765, -92.34, -90.3388235, -88.6235294, -85.1929412,
            -81.1905882, -74.3294118, -62.8941176, -45.7411765, 30.5894118, 42.3105882,
            57.1764706, 68.6117647, 74.3294118, 80.0470588, 85.1929412, 89.1952941,
            92.0541176, 94.9129412, 97.2, 97.7717647]

    fig, ax = plt.subplots(1, 1, figsize=(10, 7))
    ax.plot(current, flux, 'b-o', lw=2, markersize=4, label='Boucle d\'hystérésis (Type-96)')
    ax.set_xlabel('Courant (A)', fontsize=12)
    ax.set_ylabel('Flux (Wb-tours)', fontsize=12)
    ax.set_title("Caractéristique d'hystérésis de l'autotransformateur\n(Acier au silicium orienté Armco M4)",
                  fontsize=12, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=10)
    ax.axhline(y=0, color='k', linestyle='-', lw=0.5)
    ax.axvline(x=0, color='k', linestyle='-', lw=0.5)

    # Annotate saturation region
    ax.annotate('Zone de\nsaturation', xy=(98.2, 97.2), xytext=(60, 80),
                fontsize=10, color='red', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color='red', lw=1.5))

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task4_hysteresis.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# FIGURE 10: STC model parameters diagram (Task 3)
# ============================================================

def create_fig_task3_stc_model():
    fig, ax = plt.subplots(1, 1, figsize=(12, 6))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 7)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle STC (Saturable Transformer Component) - Schéma équivalent par phase",
                  fontsize=12, fontweight='bold')

    # Primary winding (HV - 230kV Y)
    ax.text(0.5, 5.5, 'Côté HV\n(230 kV, Y)', fontsize=9, ha='center', color='red', fontweight='bold')
    ax.plot([1.0, 2.0], [5.0, 5.0], 'r-', lw=2)

    # R_HV
    rect_rhv = FancyBboxPatch((2.0, 4.7), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_rhv)
    ax.text(2.6, 5.0, 'R_HV', fontsize=9, ha='center', va='center', fontweight='bold')

    # L_HV
    ax.plot([3.2, 3.5], [5.0, 5.0], 'k-', lw=2)
    rect_lhv = FancyBboxPatch((3.5, 4.7), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_lhv)
    ax.text(4.1, 5.0, 'L_HV', fontsize=9, ha='center', va='center', fontweight='bold')

    ax.plot([4.7, 5.5], [5.0, 5.0], 'k-', lw=2)

    # Ideal transformer
    ax.plot([5.5, 5.5], [4.0, 6.0], 'k-', lw=2)
    ax.plot([5.8, 5.8], [4.0, 6.0], 'k-', lw=2)
    ax.text(5.65, 6.3, 'Transfo\nidéal', fontsize=8, ha='center', color='purple')

    # Magnetizing branch
    ax.plot([5.65, 5.65], [4.0, 3.0], 'k-', lw=1.5)
    ax.plot([5.65, 5.0], [3.0, 3.0], 'k-', lw=1)
    ax.plot([5.65, 6.3], [3.0, 3.0], 'k-', lw=1)

    rect_rcore = FancyBboxPatch((4.5, 2.2), 1.0, 0.6, boxstyle="round,pad=0.03", facecolor='#90EE90', edgecolor='black', lw=1)
    ax.add_patch(rect_rcore)
    ax.text(5.0, 2.5, 'R_core', fontsize=8, ha='center', va='center')

    rect_lmag = FancyBboxPatch((5.8, 2.2), 1.0, 0.6, boxstyle="round,pad=0.03", facecolor='#90EE90', edgecolor='black', lw=1)
    ax.add_patch(rect_lmag)
    ax.text(6.3, 2.5, 'L_mag', fontsize=8, ha='center', va='center')

    ax.plot([5.0, 5.0], [2.8, 3.0], 'k-', lw=1)
    ax.plot([6.3, 6.3], [2.8, 3.0], 'k-', lw=1)
    ax.plot([5.0, 5.0], [1.8, 2.2], 'k-', lw=1)
    ax.plot([6.3, 6.3], [1.8, 2.2], 'k-', lw=1)
    ax.plot([5.0, 6.3], [1.8, 1.8], 'k-', lw=1)
    ax.text(5.65, 1.4, 'Pertes fer + Magnétisation', fontsize=7, ha='center', color='green', style='italic')

    # Secondary side
    ax.plot([5.8, 7.0], [5.0, 5.0], 'k-', lw=2)

    # R_LV
    rect_rlv = FancyBboxPatch((7.0, 4.7), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_rlv)
    ax.text(7.6, 5.0, 'R_LV', fontsize=9, ha='center', va='center', fontweight='bold')

    # L_LV
    ax.plot([8.2, 8.5], [5.0, 5.0], 'k-', lw=2)
    rect_llv = FancyBboxPatch((8.5, 4.7), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_llv)
    ax.text(9.1, 5.0, 'L_LV', fontsize=9, ha='center', va='center', fontweight='bold')

    ax.plot([9.7, 10.5], [5.0, 5.0], 'b-', lw=2)
    ax.text(11.0, 5.5, 'Côté LV\n(24 kV, Δ)', fontsize=9, ha='center', color='blue', fontweight='bold')

    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/task3_stc_model.png", dpi=150, bbox_inches='tight')
    plt.close()


# ============================================================
# GENERATE ALL FIGURES
# ============================================================

print("Generating figures...")
create_fig_task1_circuit()
print("  - task1_circuit.png")
create_fig_task1_params()
print("  - task1_params.png")
create_fig_task2_ct()
print("  - task2_ct_model.png")
create_fig_task2_bh()
print("  - task2_bh_curve.png")
create_fig_task3_circuit()
print("  - task3_circuit.png")
create_fig_task3_waveform()
print("  - task3_waveform.png")
create_fig_task3_stc_model()
print("  - task3_stc_model.png")
create_fig_task4_circuit()
print("  - task4_circuit.png")
create_fig_task4_inrush()
print("  - task4_inrush.png")
create_fig_task4_hysteresis()
print("  - task4_hysteresis.png")
print("All figures generated.\n")


# ============================================================
# GENERATE WORD DOCUMENT
# ============================================================

print("Generating Word document...")

doc = Document()

# ---- Styles ----
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

# ---- Title Page ----
for _ in range(4):
    doc.add_paragraph('')

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('TP N°03')
run.font.size = Pt(28)
run.bold = True
run.font.color.rgb = RGBColor(0, 51, 102)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('ANALYSE TRANSITOIRE DES TRANSFORMATEURS DE PUISSANCE')
run.font.size = Pt(16)
run.bold = True
run.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph('')
subtitle2 = doc.add_paragraph()
subtitle2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle2.add_run('Réponses détaillées')
run.font.size = Pt(14)
run.italic = True
run.font.color.rgb = RGBColor(100, 100, 100)

doc.add_paragraph('')
info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = info.add_run('TPS de Prof. A. Bayadi')
run.font.size = Pt(12)

doc.add_paragraph('')
info2 = doc.add_paragraph()
info2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = info2.add_run('Logiciel utilisé : ATP-EMTP / ATPDraw')
run.font.size = Pt(11)
run.italic = True

doc.add_page_break()

# ---- Table of Contents ----
doc.add_heading('Table des matières', level=1)
toc_items = [
    ('3.1', 'Objectifs'),
    ('3.2', 'Tâche n°1 – Transformateur monophasé saturable'),
    ('3.3', 'Tâche n°2 – Transformateur de courant (TC)'),
    ('3.4', 'Tâche n°3 – Transformateur triphasé saturable (STC)'),
    ('3.5', 'Tâche n°4 – Énergisation de l\'autotransformateur 400/132/18 kV'),
]
for num, title_text in toc_items:
    p = doc.add_paragraph()
    run = p.add_run(f'{num}  {title_text}')
    run.font.size = Pt(11)

doc.add_page_break()

# ============================================================
# 3.1 OBJECTIVES
# ============================================================

doc.add_heading('3.1. Objectifs', level=1)

p = doc.add_paragraph()
p.add_run('Les objectifs de ce TP sont les suivants :').bold = True

objectives = [
    "Se familiariser davantage avec le logiciel ATP-EMTP.",
    "Modéliser un transformateur en utilisant le modèle STC (Saturable Transformer Component) pour représenter un transformateur saturable via l'interface ATPDraw.",
    "Étudier le comportement transitoire des transformateurs de puissance.",
]
for obj in objectives:
    p = doc.add_paragraph(obj, style='List Bullet')

doc.add_paragraph('')

# ============================================================
# 3.2 TASK 1
# ============================================================

doc.add_heading('3.2. Tâche n°1 – Transformateur monophasé saturable', level=1)

doc.add_heading('Énoncé', level=2)
p = doc.add_paragraph(
    "Un transformateur triphasé 400/15,75 kV est constitué de trois unités monophasées. "
    "Il faut déterminer le modèle numérique de l'unité monophasée et simuler le court-circuit côté HT."
)

doc.add_heading('Données du problème', level=2)

headers = ['Paramètre', 'Valeur']
rows = [
    ['Puissance nominale (par phase)', 'Sn = 400/3 ≈ 133,33 MVA'],
    ['Tension de court-circuit', 'Ucc (%)'],
    ['Pertes actives (cuivre)', 'Pcc (kW)'],
    ['Pertes dans le noyau (fer)', 'Pfer (kW)'],
    ['Courant à vide', 'I0 (%)'],
    ['Rapport de transformation nominal', '400/√3 : 15,75/√3 = 230,94 : 9,093 kV'],
    ['Résistance de défaut', 'RF = 0,5 Ω'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Calcul des paramètres du transformateur', level=2)

p = doc.add_paragraph()
run = p.add_run("Hypothèse : ")
run.bold = True
p.add_run("La branche magnétisante est connectée côté HT. Les paramètres doivent être calculés pour les deux côtés HT et BT.")

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("a) Paramètres du circuit équivalent de court-circuit :")
run.bold = True

doc.add_paragraph('')

p = doc.add_paragraph()
p.add_run("Le schéma équivalent réduit au primaire d'un transformateur monophasé comprend :").italic = True

items_cc = [
    "Rcc : résistance totale de court-circuit ramenée au primaire",
    "Xcc : réactance de fuite totale de court-circuit ramenée au primaire",
    "Rm : résistance représentant les pertes fer (dans la branche magnétisante)",
    "Xm : réactance de magnétisation (dans la branche magnétisante)",
]
for item in items_cc:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("Formules de calcul :")
run.bold = True
run.underline = True

doc.add_paragraph('')

formulas = [
    "Tension nominale par phase côté HT : V1n = 400/√3 = 230,94 kV",
    "Tension nominale par phase côté BT : V2n = 15,75/√3 = 9,093 kV",
    "Courant nominal côté HT : I1n = Sn / V1n",
    "Courant nominal côté BT : I2n = Sn / V2n",
    "Rapport de transformation : a = V1n / V2n = 230,94 / 9,093 = 25,40",
    "",
    "Impédance de court-circuit (côté HT) :",
    "   Zcc = (Ucc% / 100) × V1n² / Sn",
    "   Rcc = Pcc / (3 × I1n²)  (pour les 3 phases, ou Pcc_1ph / I1n² par phase)",
    "   Xcc = √(Zcc² - Rcc²)",
    "",
    "Répartition entre les deux enroulements :",
    "   R1 = Rcc / 2   (côté HT)",
    "   R2 = Rcc / (2 × a²)   (côté BT)",
    "   L1 = Xcc / (2 × ω)   (côté HT)",
    "   L2 = Xcc / (2 × ω × a²)   (côté BT)",
    "",
    "Branche magnétisante (côté HT) :",
    "   Rm = V1n² / Pfer   (résistance pertes fer)",
    "   Xm = V1n / I0   (réactance de magnétisation)",
]

for f in formulas:
    if f == "":
        doc.add_paragraph('')
    else:
        doc.add_paragraph(f)

doc.add_paragraph('')

doc.add_heading('Schéma du circuit équivalent', level=2)
doc.add_picture(f"{IMG_DIR}/task1_circuit.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 1 : Schéma équivalent du transformateur monophasé avec court-circuit côté HT")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Diagramme des paramètres', level=2)
doc.add_picture(f"{IMG_DIR}/task1_params.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 2 : Essais de court-circuit et à vide – Détermination des paramètres")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Implémentation dans ATPDraw', level=2)

steps_atpdraw = [
    "Ouvrir ATPDraw et créer un nouveau projet.",
    "Insérer le composant « Saturable Transformer » (BCTRAN ou STC) depuis la bibliothèque de composants.",
    "Configurer les paramètres du transformateur monophasé : tensions nominales (230,94 kV / 9,093 kV), puissance nominale, pertes cuivre, pertes fer, courant à vide.",
    "Connecter la source de tension sinusoïdale côté BT (15,75/√3 kV).",
    "Ajouter l'impédance de charge Zo en conditions normales de fonctionnement.",
    "Ajouter la résistance de défaut RF = 0,5 Ω avec un interrupteur commandé côté HT pour simuler le court-circuit.",
    "Configurer les paramètres de simulation : pas de calcul (Δt), durée de simulation, instant de fermeture du court-circuit.",
    "Lancer la simulation et observer les courants et tensions transitoires.",
]
for i, step in enumerate(steps_atpdraw, 1):
    doc.add_paragraph(f"{i}. {step}")

doc.add_paragraph('')

doc.add_heading('Résultats attendus et discussion', level=2)

p = doc.add_paragraph()
p.add_run("Lors de la simulation du court-circuit côté HT :").bold = True

results_task1 = [
    "Le courant de court-circuit présente une composante transitoire (apériodique DC) superposée à la composante en régime permanent (AC).",
    "La composante DC décroît exponentiellement avec une constante de temps τ = L/R, où L et R sont l'inductance et la résistance totales du circuit de court-circuit.",
    "L'amplitude maximale du courant de court-circuit dépend de l'instant de fermeture du défaut par rapport à la forme d'onde de la tension.",
    "Le courant asymétrique maximal (valeur crête) peut atteindre environ 2,5 à 2,8 fois le courant de court-circuit symétrique en régime permanent.",
    "La résistance RF = 0,5 Ω limite l'amplitude du courant de court-circuit par rapport à un court-circuit franc.",
    "En régime permanent de court-circuit, le courant est limité par l'impédance de court-circuit Zcc du transformateur plus RF.",
]
for r in results_task1:
    doc.add_paragraph(r, style='List Bullet')

doc.add_page_break()

# ============================================================
# 3.3 TASK 2
# ============================================================

doc.add_heading('3.3. Tâche n°2 – Transformateur de courant (TC)', level=1)

doc.add_heading('Énoncé', level=2)
p = doc.add_paragraph(
    "Modéliser un transformateur de courant (TC) de type 5P20 dans ATPDraw en utilisant "
    "l'élément pseudo-non-linéaire de type 98 pour représenter la caractéristique magnétisante, "
    "avec possibilité de fixer la valeur initiale du flux."
)

doc.add_heading('Données du TC 5P20', level=2)

headers = ['Paramètre', 'Valeur']
rows = [
    ['Puissance nominale', 'Sn = 20 VA'],
    ['Impédance de charge nominale', 'Zobc_n = 20 Ω, cos φ = 0,5'],
    ['Rapport de courant', '500:1 (I1n:I2n) A/A'],
    ['Rapport de spires', '500:1 (N2:N1)'],
    ['Section du noyau', 'S = 28,8 cm² = 2,88 × 10⁻³ m²'],
    ['Longueur du circuit magnétique', 'l = 0,675 m'],
    ['Résistance secondaire', 'R2 = 4,5 Ω'],
    ['Charge nominale du TC', 'Zobc = 10 + j17,3 Ω'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Calcul des paramètres du TC', level=2)

p = doc.add_paragraph()
run = p.add_run("a) Courants nominaux :")
run.bold = True

calcs_ct = [
    "I2n = √(Sn / Zobc_n) = √(20 / 20) = 1 A",
    "I1n = 500 A (rapport 500:1)",
    "",
    "b) Courant d'excitation pour le facteur de surcourant 20 :",
    "Le courant primaire maximal : I1_max = 20 × I1n = 20 × 500 = 10 000 A",
    "Le courant secondaire correspondant : I2_max = 10 000 / 500 = 20 A",
    "",
    "c) Transformation de la caractéristique B-H en λ = f(iμ) :",
    "Les formules de conversion sont :",
    "   iμ = (l / N2²) × H = (0,675 / 500) × H = 1,35 × 10⁻³ × H",
    "   λ = N2 × S × B = 500 × 2,88 × 10⁻³ × B = 1,44 × B",
]
for c in calcs_ct:
    if c == "":
        doc.add_paragraph('')
    else:
        doc.add_paragraph(c)

doc.add_paragraph('')

doc.add_heading('Caractéristique magnétisante λ = f(iμ)', level=2)

headers = ['iμ (A)', 'λ (Vs)', 'B (T)', 'H (A/m)']
i_data = [0.0143, 0.0382, 0.0573, 0.0955, 0.1909, 0.7637, 3.8184, 28.6378]
lambda_data = [0.1440, 1.4400, 2.0160, 2.3184, 2.4912, 2.6928, 2.8368, 2.9664]
N2 = 500
S_core = 2.88e-3
l_core = 0.675
B_data = [lam / (N2 * S_core) for lam in lambda_data]
H_data = [i * N2 / l_core for i in i_data]

rows = []
for i in range(len(i_data)):
    rows.append([f'{i_data[i]:.4f}', f'{lambda_data[i]:.4f}', f'{B_data[i]:.4f}', f'{H_data[i]:.2f}'])

add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/task2_bh_curve.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 3 : Caractéristique magnétisante λ = f(iμ) et courbe B-H du TC")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Schéma du modèle du TC dans ATPDraw', level=2)

doc.add_picture(f"{IMG_DIR}/task2_ct_model.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 4 : Circuit du modèle du TC dans ATPDraw")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Description du modèle', level=2)

p = doc.add_paragraph()
p.add_run("Le modèle du TC dans ATPDraw comprend les éléments suivants :").bold = True

elements_ct = [
    "Source de courant : une source sinusoïdale qui génère le courant primaire. Pour obtenir un rapport de surcourant de 20, le courant de source est I = 20 × 500 = 10 000 A (valeur crête : 10 000 × √2 ≈ 14 142 A).",
    "Impédance de source (Zs) : représente l'impédance de la source côté primaire.",
    "Transformateur idéal : avec un rapport de spires N2:N1 = 500:1. Dans ATP-EMTP, le TC est modélisé comme un transformateur de tension inversé (le secondaire a plus de spires).",
    "Résistance secondaire R2 = 4,5 Ω : résistance de l'enroulement secondaire.",
    "Branche magnétisante : composée de l'élément non-linéaire Type 98, qui représente la caractéristique de saturation λ = f(iμ), connecté en série avec une source de tension DC (Type 11) pour fixer la valeur initiale du flux.",
    "Charge du TC : Zobc = 10 + j17,3 Ω (impédance complexe représentant la charge nominale).",
]
for e in elements_ct:
    doc.add_paragraph(e, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Paramètres de simulation', level=2)

sim_params = [
    "Pas de calcul : ΔT = 10⁻⁵ s (10 μs)",
    "Ce pas de calcul est suffisamment petit pour capturer les harmoniques générées par la saturation du noyau.",
    "La fréquence d'échantillonnage correspondante est de 100 kHz.",
]
for s in sim_params:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Résultats attendus et discussion', level=2)

results_ct = [
    "En fonctionnement normal (sans saturation), le courant secondaire reproduit fidèlement le courant primaire avec un rapport 500:1.",
    "Lorsque le courant primaire atteint des valeurs élevées (facteur de surcourant = 20), le noyau entre en saturation.",
    "La saturation du noyau provoque une distorsion du courant secondaire : il ne suit plus fidèlement le courant primaire.",
    "Le courant magnétisant augmente fortement dans la zone de saturation, ce qui dévie une partie du courant du secondaire vers la branche magnétisante.",
    "L'erreur de transformation augmente avec le degré de saturation.",
    "La classe 5P20 signifie que l'erreur composite ne dépasse pas 5% pour un facteur de surcourant de 20.",
    "La présence d'une composante DC (par exemple lors d'un court-circuit asymétrique) peut provoquer une saturation unilatérale rapide du noyau, dégradant les performances du TC.",
]
for r in results_ct:
    doc.add_paragraph(r, style='List Bullet')

doc.add_page_break()

# ============================================================
# 3.4 TASK 3
# ============================================================

doc.add_heading('3.4. Tâche n°3 – Transformateur triphasé saturable (Modèle STC)', level=1)

doc.add_heading('Énoncé', level=2)
p = doc.add_paragraph(
    "Modéliser un transformateur triphasé 24 kV : 230 kV, 150 MVA, couplage Δ/Y-mis à la terre, "
    "en utilisant le modèle STC dans ATPDraw, puis simuler un court-circuit triphasé au secondaire."
)

doc.add_heading('Données du transformateur', level=2)

headers = ['Paramètre', 'Valeur']
rows = [
    ['Puissance nominale', '150 MVA'],
    ['Tensions nominales', '24 kV (Δ) / 230 kV (Y-gnd)'],
    ['Rapport X/R côté 24 kV', '10'],
    ['Réactance de fuite côté 24 kV', '5,5 %'],
    ['Rapport X/R côté 230 kV', '12'],
    ['Réactance de fuite côté 230 kV', '7 %'],
    ['Pertes dans le noyau', '50 kW par phase (côté Y)'],
    ['Inductance de magnétisation', '20 kΩ à tension nominale (côté HT)'],
    ['Fréquence', '60 Hz'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Calcul des paramètres du modèle STC', level=2)

p = doc.add_paragraph()
run = p.add_run("a) Impédance de base :")
run.bold = True

calcs_stc = [
    "",
    "Côté HT (230 kV, Y) :",
    "   V_base_HV = 230 / √3 = 132,79 kV (par phase)",
    "   I_base_HV = 150 × 10⁶ / (√3 × 230 × 10³) = 376,5 A",
    "   Z_base_HV = V_base_HV / I_base_HV = 132 790 / 376,5 = 352,7 Ω",
    "",
    "Côté BT (24 kV, Δ) :",
    "   V_base_LV = 24 kV (tension de ligne = tension de phase en delta)",
    "   I_base_LV = 150 × 10⁶ / (√3 × 24 × 10³) = 3 608 A (ligne)",
    "   I_phase_LV = 3 608 / √3 = 2 083 A (par phase en delta)",
    "   Z_base_LV = V_base_LV² / (Sn/3) = (24 000)² / (50 × 10⁶) = 11,52 Ω",
    "",
]
for c in calcs_stc:
    if c == "":
        doc.add_paragraph('')
    else:
        doc.add_paragraph(c)

p = doc.add_paragraph()
run = p.add_run("b) Paramètres de l'enroulement HT (230 kV, Y) :")
run.bold = True

calcs_hv = [
    "   X_HV = 7% × Z_base_HV = 0,07 × 352,7 = 24,69 Ω",
    "   R_HV = X_HV / (X/R) = 24,69 / 12 = 2,058 Ω",
    "   L_HV = X_HV / (2π × 60) = 24,69 / 377 = 0,0655 H = 65,5 mH",
]
for c in calcs_hv:
    doc.add_paragraph(c)

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("c) Paramètres de l'enroulement BT (24 kV, Δ) :")
run.bold = True

calcs_lv = [
    "   X_LV = 5,5% × Z_base_LV = 0,055 × 11,52 = 0,634 Ω",
    "   R_LV = X_LV / (X/R) = 0,634 / 10 = 0,0634 Ω",
    "   L_LV = X_LV / (2π × 60) = 0,634 / 377 = 1,68 mH",
]
for c in calcs_lv:
    doc.add_paragraph(c)

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("d) Branche magnétisante (côté HT) :")
run.bold = True

calcs_mag = [
    "   Pertes fer par phase : Pfer = 50 kW",
    "   Tension par phase côté HT : V_HV = 230/√3 = 132,79 kV",
    "   R_core = V_HV² / Pfer = (132 790)² / (50 × 10³) = 352 619 Ω ≈ 352,6 kΩ",
    "   L_mag = 20 kΩ / (2π × 60) = 20 000 / 377 = 53,05 H",
    "   (Ou directement : X_mag = 20 kΩ, donc L_mag = X_mag / ω = 53,05 H)",
]
for c in calcs_mag:
    doc.add_paragraph(c)

doc.add_paragraph('')

doc.add_heading('Schéma du modèle STC par phase', level=2)
doc.add_picture(f"{IMG_DIR}/task3_stc_model.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 5 : Modèle STC par phase du transformateur")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Tableau récapitulatif des paramètres STC', level=2)

headers = ['Paramètre', 'Valeur', 'Côté']
rows = [
    ['R_HV', '2,058 Ω', 'HT (230 kV)'],
    ['L_HV (X_HV)', '65,5 mH (24,69 Ω)', 'HT (230 kV)'],
    ['R_LV', '0,0634 Ω', 'BT (24 kV)'],
    ['L_LV (X_LV)', '1,68 mH (0,634 Ω)', 'BT (24 kV)'],
    ['R_core', '352,6 kΩ', 'Magnétisation'],
    ['L_mag (X_mag)', '53,05 H (20 kΩ)', 'Magnétisation'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Question 1 : Implémentation dans ATPDraw', level=2)

steps_task3 = [
    "Ouvrir ATPDraw et sélectionner le composant « Saturable Transformer Component (STC) » triphasé.",
    "Configurer le couplage : Delta (Δ) côté 24 kV, Y-mis à la terre (Yn) côté 230 kV.",
    "Entrer les paramètres calculés ci-dessus pour chaque enroulement.",
    "Configurer la branche magnétisante avec R_core et L_mag.",
    "Connecter la source triphasée 230 kV, 60 Hz côté HT.",
    "Ajouter un interrupteur triphasé (disjoncteur S1) pour le court-circuit au secondaire (côté 24 kV), commandé à t = 0,1 s.",
    "Configurer la durée de simulation à 0,3 s avec un pas de calcul approprié (ex. : ΔT = 10 μs).",
]
for i, s in enumerate(steps_task3, 1):
    doc.add_paragraph(f"{i}. {s}")

doc.add_paragraph('')

doc.add_heading('Schéma du circuit dans ATPDraw', level=2)
doc.add_picture(f"{IMG_DIR}/task3_circuit.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 6 : Schéma ATPDraw du circuit avec court-circuit triphasé au secondaire")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Question 2 : Simulation du court-circuit triphasé', level=2)

p = doc.add_paragraph()
p.add_run("Conditions de simulation :").bold = True

sim_cond = [
    "Source triphasée : 230 kV (ligne-ligne), 60 Hz",
    "Court-circuit triphasé au secondaire (24 kV) à t = 0,1 s",
    "Durée de simulation : 0,3 s",
    "Avant le court-circuit (0 < t < 0,1 s) : le transformateur fonctionne à vide ou sous charge nominale",
    "Après le court-circuit (0,1 s < t < 0,3 s) : courant de court-circuit transitoire",
]
for s in sim_cond:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 3 : Courant dans le disjoncteur S1 et valeurs RMS', level=2)

doc.add_picture(f"{IMG_DIR}/task3_waveform.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 7 : Courant dans le disjoncteur S1 et valeur RMS (simulation illustrative)")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("Analyse du courant de court-circuit :")
run.bold = True

analysis_sc = [
    "Le courant nominal côté HT (230 kV) : I_n = 150 × 10⁶ / (√3 × 230 × 10³) ≈ 376,5 A",
    "L'impédance totale de court-circuit (% total) ≈ 5,5% + 7% = 12,5%",
    "Le courant de court-circuit symétrique : I_cc = I_n / (Xcc_total%) = 376,5 / 0,125 ≈ 3 012 A (RMS)",
    "La valeur crête asymétrique (avec composante DC) : I_peak ≈ 2,55 × I_cc ≈ 7 681 A",
    "La constante de temps de décroissance DC : τ = X/(R×ω) = L_total / R_total",
    "Avec les paramètres calculés : τ ≈ (X_HV + X_LV×a²) / (R_HV + R_LV×a²) / ω",
    "La valeur RMS du courant atteint rapidement sa valeur en régime permanent après quelques cycles.",
]
for a in analysis_sc:
    doc.add_paragraph(a, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 4 : Discussion des résultats', level=2)

discussion_task3 = [
    "1. Composante transitoire DC : Au moment du court-circuit (t = 0,1 s), un offset DC apparaît dans le courant. "
    "Son amplitude dépend de l'instant de fermeture par rapport au zéro de la tension. Cette composante décroît "
    "exponentiellement avec la constante de temps τ = L/R du circuit.",

    "2. Asymétrie entre phases : Les trois phases ne présentent pas le même comportement transitoire car le "
    "court-circuit se produit à un instant donné du cycle. La phase dont la tension est au voisinage de zéro "
    "au moment du défaut présente la plus grande asymétrie (composante DC maximale).",

    "3. Saturation du noyau : Le modèle STC prend en compte la saturation magnétique. Lors de la surintensité "
    "de court-circuit, si le flux dans le noyau dépasse le flux de saturation, le courant magnétisant augmente "
    "fortement. Cela peut amplifier le courant total.",

    "4. Valeur RMS : La valeur RMS du courant de court-circuit est importante pour le dimensionnement des "
    "équipements de protection. La valeur maximale du courant asymétrique (première crête) détermine "
    "la tenue électrodynamique requise pour le disjoncteur.",

    "5. Influence du couplage Δ-Yn : Le couplage delta côté BT bloque les composantes homopolaires "
    "(harmonique 3 et ses multiples) dans le delta. Le neutre mis à la terre côté Y permet l'écoulement "
    "des courants de défaut à la terre.",

    "6. Décroissance vers le régime permanent : Après plusieurs constantes de temps (5τ), la composante DC "
    "a pratiquement disparu et le courant de court-circuit est purement sinusoïdal avec une amplitude "
    "déterminée par l'impédance de court-circuit du transformateur.",
]
for d in discussion_task3:
    p = doc.add_paragraph(d)

doc.add_page_break()

# ============================================================
# 3.5 TASK 4
# ============================================================

doc.add_heading('3.5. Tâche n°4 – Énergisation de l\'autotransformateur 400/132/18 kV', level=1)

doc.add_heading('Énoncé', level=2)
p = doc.add_paragraph(
    "Étudier l'énergisation d'un autotransformateur triphasé à trois enroulements couplé Yyd "
    "(400/132/18 kV). Les enroulements 132 kV (Y) et 18 kV (delta) sont à vide. "
    "La branche magnétisante non-linéaire est représentée par des inducteurs hystérétiques Type-96."
)

doc.add_heading('Données de la plaque signalétique', level=2)

headers = ['Paramètre', 'Valeur']
rows = [
    ['Tensions nominales', '400/132/18 kV'],
    ['Couplage', 'Yyd (Y - y - delta)'],
    ['Enroulements 132 kV et 18 kV', 'À vide (non chargés)'],
    ['Branche magnétisante', 'Inducteurs hystérétiques Type-96, couplés en delta'],
    ['Matériau du noyau', 'Acier au silicium orienté Armco M4'],
    ['Source', '420 kV, 60 Hz, triphasée'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Séquence de manœuvres', level=2)

headers = ['Événement', 'Instant', 'Description']
rows = [
    ['Désenergisation', 't = 45 ms', 'Ouverture du disjoncteur, coupure de l\'alimentation'],
    ['Réenergisation Phase A', 't = 73,5 ms', 'Fermeture du disjoncteur pour la phase A'],
    ['Réenergisation Phases B, C', 't = 78,5 ms', 'Fermeture du disjoncteur pour les phases B et C'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

p = doc.add_paragraph()
p.add_run("Durée de simulation : 0,15 s (150 ms)").bold = True

doc.add_paragraph('')

doc.add_heading('Caractéristique d\'hystérésis du noyau', level=2)

doc.add_picture(f"{IMG_DIR}/task4_hysteresis.png", width=Inches(5.5))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 8 : Caractéristique d'hystérésis de l'autotransformateur (Type-96, Armco M4)")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

p = doc.add_paragraph()
p.add_run("Données de la courbe d'hystérésis (fournies dans le fichier ATP) :").bold = True

doc.add_paragraph('')

headers = ['Courant (A)', 'Flux (Wb-tours)']
current_data = [-36.825, -24.55, -11.0475, -4.91, -1.84125, 0.61375, 2.148125,
                3.55975, 4.29625, 4.91, 6.1375, 6.75125, 8.5925, 11.0475,
                13.37975, 17.491875, 23.93625, 32.835625, 42.9625, 61.375, 98.2, 135.025]
flux_data = [-94.9129, -94.3412, -92.34, -90.3388, -88.6235, -85.1929,
             -81.1906, -74.3294, -62.8941, -45.7412, 30.5894, 42.3106,
             57.1765, 68.6118, 74.3294, 80.0471, 85.1929, 89.1953,
             92.0541, 94.9129, 97.2, 97.7718]

rows = []
for i in range(len(current_data)):
    rows.append([f'{current_data[i]:.4f}', f'{flux_data[i]:.4f}'])
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Question 1 : Implémentation dans ATPDraw', level=2)

steps_task4 = [
    "Créer un nouveau projet dans ATPDraw.",
    "Insérer la source triphasée 420 kV, 60 Hz avec les impédances de source Zs.",
    "Configurer les interrupteurs triphasés avec les temps de manœuvre spécifiés :\n"
    "   - Ouverture : t = 45 ms (pour les 3 phases)\n"
    "   - Fermeture : t = 73,5 ms (phase A), t = 78,5 ms (phases B et C)",
    "Insérer le modèle d'autotransformateur avec couplage Yyd.",
    "Configurer la branche magnétisante non-linéaire avec les inducteurs hystérétiques Type-96 couplés en delta.",
    "Entrer les données de la courbe d'hystérésis (22 points courant-flux).",
    "Les enroulements 132 kV et 18 kV sont laissés en circuit ouvert (à vide).",
    "Configurer la durée de simulation à 0,15 s avec un pas de calcul ΔT approprié.",
]
for i, s in enumerate(steps_task4, 1):
    doc.add_paragraph(f"{i}. {s}")

doc.add_paragraph('')

doc.add_heading('Schéma du circuit', level=2)

doc.add_picture(f"{IMG_DIR}/task4_circuit.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 9 : Schéma du circuit d'énergisation de l'autotransformateur dans ATPDraw")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Question 2 : Simulation', level=2)

p = doc.add_paragraph()
p.add_run("Paramètres de simulation :").bold = True

sim_params_4 = [
    "Source : 420 kV (ligne-ligne), 60 Hz",
    "Durée de simulation : 150 ms",
    "Désenergisation à t = 45 ms",
    "Réenergisation : phase A à t = 73,5 ms ; phases B et C à t = 78,5 ms",
    "Pas de calcul : ΔT suffisamment petit pour capturer les phénomènes transitoires (ex. 10⁻⁵ s)",
]
for s in sim_params_4:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 3 : Courant magnétisant en régime permanent', level=2)

p = doc.add_paragraph(
    "En régime permanent (avant la désenergisation, 0 < t < 45 ms), le courant magnétisant "
    "du transformateur est très faible par rapport au courant nominal. Il est typiquement de l'ordre "
    "de 0,5% à 2% du courant nominal."
)

regime_permanent = [
    "Le courant magnétisant est quasi-sinusoïdal si le transformateur fonctionne dans la zone linéaire de la courbe B-H.",
    "Il contient des harmoniques (principalement l'harmonique 3) dues à la non-linéarité de la caractéristique de magnétisation.",
    "La forme d'onde est légèrement distordue, avec des pics correspondant aux passages dans la zone de saturation.",
    "L'amplitude crête du courant magnétisant est typiquement quelques ampères pour un transformateur de cette puissance.",
    "Les pertes fer (pertes par hystérésis et par courants de Foucault) sont constantes à tension nominale.",
]
for r in regime_permanent:
    doc.add_paragraph(r, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 4 : Courant d\'appel (Inrush) lors de la réenergisation', level=2)

doc.add_picture(f"{IMG_DIR}/task4_inrush.png", width=Inches(6.0))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 10 : Courant magnétisant et courant d'appel lors de la réenergisation")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

p = doc.add_paragraph()
run = p.add_run("Phénomène du courant d'appel (inrush current) :")
run.bold = True

doc.add_paragraph('')

inrush_desc = [
    "Lors de la réenergisation d'un transformateur, un courant transitoire très élevé appelé « courant d'appel » "
    "(inrush current) peut se produire. Ce phénomène est dû à la saturation du noyau magnétique.",

    "Le courant d'appel dépend de :",
    "   • L'instant de fermeture du disjoncteur par rapport au zéro de tension",
    "   • Le flux rémanent dans le noyau au moment de la réenergisation",
    "   • Les caractéristiques non-linéaires du noyau (courbe d'hystérésis)",
    "   • L'impédance de la source",
    "",
    "Caractéristiques du courant d'appel observé :",
]
for d in inrush_desc:
    if d == "":
        doc.add_paragraph('')
    else:
        doc.add_paragraph(d)

inrush_chars = [
    "Le courant d'appel peut atteindre 5 à 10 fois le courant nominal du transformateur.",
    "Il est fortement asymétrique (composante DC importante) et contient de nombreuses harmoniques, "
    "en particulier l'harmonique 2.",
    "La forme d'onde est caractéristique : des impulsions unipolaires avec de longues périodes à zéro.",
    "La décroissance est lente, de l'ordre de plusieurs secondes pour les grands transformateurs, "
    "avec une constante de temps déterminée par le rapport L_mag/R_total.",
    "La fermeture décalée des phases (A à 73,5 ms, B et C à 78,5 ms) produit des courants d'appel "
    "différents pour chaque phase.",
    "La phase A, réenergisée en premier, peut avoir un courant d'appel plus important si l'instant "
    "de fermeture coïncide avec un zéro de tension.",
]
for c in inrush_chars:
    doc.add_paragraph(c, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 5 : Discussion des résultats', level=2)

discussion_task4 = [
    "1. Flux rémanent : Lors de la désenergisation à t = 45 ms, le flux dans le noyau ne retombe pas à zéro "
    "mais conserve une valeur rémanente déterminée par la courbe d'hystérésis. Cette valeur de flux rémanent "
    "dépend du point de fonctionnement au moment de l'ouverture du disjoncteur.",

    "2. Saturation profonde : Lors de la réenergisation, si le flux initial (rémanent) et le flux imposé par "
    "la tension de la source sont de même signe, le flux total peut atteindre des valeurs très élevées, "
    "poussant le noyau en saturation profonde. Cela provoque une augmentation dramatique du courant magnétisant.",

    "3. Rôle du Type-96 (hystérésis) : Le modèle Type-96 permet de capturer fidèlement le comportement "
    "hystérétique du noyau, y compris le flux rémanent après désenergisation. C'est essentiel pour une "
    "simulation réaliste du courant d'appel. Un modèle sans hystérésis (Type-98 seul) sous-estimerait "
    "le flux rémanent.",

    "4. Impact de l'instant de fermeture : La fermeture non simultanée des trois phases (phase A à 73,5 ms, "
    "phases B et C à 78,5 ms) simule le comportement réel d'un disjoncteur. Cette stratégie peut être "
    "optimisée (fermeture contrôlée ou « point-on-wave switching ») pour minimiser le courant d'appel.",

    "5. Couplage Yyd : Le couplage delta du troisième enroulement (18 kV) fournit un chemin pour les "
    "courants harmoniques de rang 3 et leurs multiples, même à vide. Cela affecte la forme d'onde du "
    "courant magnétisant et du courant d'appel.",

    "6. Protection : Le courant d'appel est un défi pour les protections différentielles du transformateur. "
    "La présence importante de l'harmonique 2 dans le courant d'appel est utilisée comme critère de "
    "blocage de la protection différentielle pour éviter les déclenchements intempestifs lors de "
    "l'énergisation.",

    "7. Comparaison régime permanent / transitoire : Le courant magnétisant en régime permanent est "
    "typiquement inférieur à 1-2% du courant nominal, tandis que le courant d'appel peut atteindre "
    "5 à 10 fois ce courant nominal. Cette différence illustre l'importance de la modélisation "
    "non-linéaire pour l'étude des phénomènes transitoires dans les transformateurs.",
]
for d in discussion_task4:
    p = doc.add_paragraph(d)

doc.add_paragraph('')
doc.add_paragraph('')

# ---- Final note ----
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("— Fin du document —")
run.italic = True
run.font.color.rgb = RGBColor(128, 128, 128)

# Save
output_path = "/workspace/TP03_Reponses.docx"
doc.save(output_path)
print(f"Document saved: {output_path}")
print("Done!")
