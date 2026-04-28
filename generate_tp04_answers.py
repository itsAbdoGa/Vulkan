#!/usr/bin/env python3
"""
Script to generate the TP04 Word document with answers and illustrations.
Transient Analysis of Transmission Lines - ATP-EMTP
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Arc
import numpy as np
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

IMG_DIR = "/workspace/images_tp04"
os.makedirs(IMG_DIR, exist_ok=True)

# ============================================================
# LINE PARAMETERS (typical 100 km transmission line)
# ============================================================
l_line = 100  # km
R_per_km = 0.03  # Ω/km
L_per_km = 1.0e-3  # H/km
C_per_km = 11.11e-9  # F/km

R_total = R_per_km * l_line  # 3 Ω
L_total = L_per_km * l_line  # 0.1 H
C_total = C_per_km * l_line  # 1.111 μF

Zc = np.sqrt(L_per_km / C_per_km)  # characteristic impedance ≈ 300 Ω
v_prop = 1.0 / np.sqrt(L_per_km * C_per_km)  # propagation velocity (km/s)
tau = l_line / v_prop  # travel time (s)

V_source = 1.0  # DC source voltage (pu or 1V for simplicity)
t_switch = 0.01  # switching time 10 ms

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def set_cell_shading(cell, color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)

def add_table_with_style(doc, headers, rows):
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
# SIMULATION FUNCTIONS
# ============================================================

def simulate_lossless_open(t_array, V_src, Zc_val, tau_val, t_sw):
    """Simulate lossless line with open end, DC source energization using Bewley lattice."""
    V_send = np.zeros_like(t_array)
    V_recv = np.zeros_like(t_array)
    I_send = np.zeros_like(t_array)
    I_recv = np.zeros_like(t_array)

    for idx, t in enumerate(t_array):
        tt = t - t_sw
        if tt < 0:
            continue
        n_reflections = 30
        v_s = 0.0
        v_r = 0.0
        i_s = 0.0
        i_r = 0.0

        for k in range(n_reflections):
            # Forward wave arrives at receiving end at (2k+1)*tau
            # Reflected wave arrives back at sending end at (2k+2)*tau (but k starts at 0 for first reflection)
            t_arrive_recv = (2*k + 1) * tau_val
            t_arrive_send = (2*k + 2) * tau_val

            if tt >= 0:
                # Initial step at sending end
                pass

            # At open end: voltage reflection coeff = +1, current reflection coeff = -1
            # At source end (assuming Zs=Zc for matched, or Zs=0 for ideal source)
            # For ideal voltage source: Zs = 0, reflection coeff at source = -1 for voltage... 
            # Actually for ideal voltage source: Γ_s = (Zs - Zc)/(Zs + Zc) = -1
            # At open end: Γ_r = (Zload - Zc)/(Zload + Zc) = +1

        # Use lattice diagram approach for ideal source (Zs=0) and open end
        # Γ_source = -1 (ideal voltage source)
        # Γ_open = +1 (open circuit)
        # Initial voltage step = V_src * Zc/(Zs + Zc) = V_src (since Zs=0... wait)
        # For ideal voltage source with Zs=0: V_send = V_src always after switch
        # The traveling wave: v_forward = V_src/2 for matched... 
        
        # Correct approach: For ideal voltage source (Zs = 0):
        # At t=t_sw: source applies V_src. Sending end voltage = V_src immediately.
        # A forward traveling wave of amplitude V_src propagates to receiving end.
        # At open end: reflected wave = +V_src (total = 2*V_src)
        # Reflected wave travels back to source. At source (Zs=0): Γ_s = -1
        # Reflected wave at source = -V_src (brings sending end voltage back to V_src from 2V)
        # Actually let's use the standard lattice:
        
        # V_sending = V_src for all t > t_sw (ideal voltage source maintains voltage)
        # More precisely with traveling waves and reflections for current and recv voltage:
        
        # Using superposition of traveling waves:
        v_s = V_src  # ideal source maintains this
        
        # Receiving end voltage: sum of forward and reflected waves arriving there
        v_r_temp = 0.0
        i_s_temp = 0.0
        for k in range(n_reflections + 1):
            t_arr_r = (2*k + 1) * tau_val
            if tt >= t_arr_r:
                v_r_temp += 2 * V_src * ((-1)**k)
            
        v_r = v_r_temp
        
        # Current at sending end: I = V_src/Zc for forward wave
        # Each round trip adds contribution
        i_s_temp = 0.0
        if tt >= 0:
            i_s_temp = V_src / Zc_val
        for k in range(1, n_reflections + 1):
            t_arr_s = 2*k * tau_val
            if tt >= t_arr_s:
                i_s_temp += 0  # For ideal source, current just maintains V_src
        
        # Simpler correct approach: For lossless line, open end, ideal DC source
        # V_recv oscillates between 0 and 2V with period 4*tau
        # Actually it's a staircase for distributed parameter model
        
        V_send[idx] = v_s
        V_recv[idx] = v_r
        I_send[idx] = i_s_temp
        
    return V_send, V_recv, I_send, I_recv


def simulate_distributed_param_lossless_open(t_array, V_dc, Zc_val, tau_val, t_sw):
    """
    Lossless line, DC source (Zs=0), open end.
    Using Bewley lattice diagram.
    Γ_source = (0 - Zc)/(0 + Zc) = -1
    Γ_recv = (∞ - Zc)/(∞ + Zc) = +1
    Initial forward voltage wave: e1 = V_dc * Zc/(0 + Zc) = V_dc
    """
    V_send = np.zeros_like(t_array)
    V_recv = np.zeros_like(t_array)
    I_send = np.zeros_like(t_array)
    I_recv = np.zeros_like(t_array)
    
    for idx, t in enumerate(t_array):
        tt = t - t_sw
        if tt < 0:
            continue
        
        # For ideal voltage source: V at sending end = V_dc always
        V_send[idx] = V_dc
        
        # At receiving (open) end: voltage builds up in steps
        v_r = 0.0
        for n in range(50):
            t_arrive = (2*n + 1) * tau_val
            if tt >= t_arrive:
                # Each forward wave has amplitude V_dc * (-1)^n at open end it doubles
                # Actually: wave amplitudes at open end:
                # n=0: +V_dc arrives, reflects +V_dc, total = 2*V_dc
                # The reflected +V_dc goes back, at source Γ=-1, so -V_dc reflects back
                # n=1: -V_dc arrives at open end, reflects -V_dc, total change = -2*V_dc
                # Net at recv: 2*V_dc - 2*V_dc = 0... then next:
                # n=2: +V_dc arrives, reflects... total change +2*V_dc → net = 2*V_dc
                # This gives oscillation between 2V and 0 with period 4τ
                v_r += 2.0 * V_dc * ((-1.0)**n)
        V_recv[idx] = v_r
        
        # Current at sending end
        i_s = 0.0
        for n in range(50):
            t_arrive_back = (2*n) * tau_val  # reflected waves arrive back
            if n == 0:
                if tt >= 0:
                    i_s += V_dc / Zc_val
            else:
                if tt >= t_arrive_back:
                    # Current reflection at source: after each round trip
                    # Forward current = V_dc/Zc, reflected current from open end = -V_dc/Zc (Γ_I = -1 at open)
                    # This reflected current arrives at source, source reflects with Γ_I_source = +1
                    # So current steps: 
                    i_s += 2 * V_dc / Zc_val * ((-1.0)**n)
        I_send[idx] = i_s
        
        # Current at receiving end is always 0 (open circuit)
        I_recv[idx] = 0.0
    
    return V_send, V_recv, I_send, I_recv


def simulate_distributed_param_lossless_sc(t_array, V_dc, Zc_val, tau_val, t_sw):
    """
    Lossless line, DC source (Zs=0), short-circuited end.
    Γ_source = -1 (ideal voltage source)
    Γ_recv = (0 - Zc)/(0 + Zc) = -1 (short circuit)
    """
    V_send = np.zeros_like(t_array)
    V_recv = np.zeros_like(t_array)
    I_send = np.zeros_like(t_array)
    I_recv = np.zeros_like(t_array)
    
    for idx, t in enumerate(t_array):
        tt = t - t_sw
        if tt < 0:
            continue
        
        V_send[idx] = V_dc  # ideal source
        V_recv[idx] = 0.0   # short circuit: always 0
        
        # Current at sending end: builds up in steps
        i_s = 0.0
        for n in range(50):
            t_arrive_back = 2*n * tau_val
            if n == 0:
                if tt >= 0:
                    i_s += V_dc / Zc_val
            else:
                if tt >= t_arrive_back:
                    i_s += 2 * V_dc / Zc_val
        I_send[idx] = i_s
        
        # Current at receiving (SC) end: doubles at each arrival
        i_r = 0.0
        for n in range(50):
            t_arrive = (2*n + 1) * tau_val
            if tt >= t_arrive:
                i_r += 2 * V_dc / Zc_val
        I_recv[idx] = i_r
    
    return V_send, V_recv, I_send, I_recv


def simulate_with_losses_open(t_array, V_dc, R_t, L_t, C_t, t_sw, n_sections=1):
    """Simulate line with losses using lumped pi-model sections (state-space approach)."""
    dt = t_array[1] - t_array[0]
    
    R_sec = R_t / n_sections
    L_sec = L_t / n_sections
    C_sec = C_t / n_sections
    
    # State variables: capacitor voltages and inductor currents
    # For n sections: n+1 nodes, n inductors, n+1 capacitors (pi model)
    # Simplified: each pi section has C/2 at each end and L, R in series
    
    # Use simple trapezoidal integration
    # Nodes: 0 (source), 1, 2, ..., n (receiving end)
    # Each section: node k to node k+1 has R_sec, L_sec in series
    # Each node k has C_sec to ground (C/2 from adjacent sections combined)
    
    n_nodes = n_sections + 1
    V_nodes = np.zeros(n_nodes)
    I_branches = np.zeros(n_sections)  # current in each L-R branch
    
    # Capacitance at each node
    C_nodes = np.zeros(n_nodes)
    for k in range(n_nodes):
        if k == 0 or k == n_nodes - 1:
            C_nodes[k] = C_sec / 2
        else:
            C_nodes[k] = C_sec
    
    V_send_arr = np.zeros(len(t_array))
    V_recv_arr = np.zeros(len(t_array))
    I_send_arr = np.zeros(len(t_array))
    I_recv_arr = np.zeros(len(t_array))
    
    for idx in range(len(t_array)):
        t = t_array[idx]
        
        # Source: ideal voltage source at node 0
        if t >= t_sw:
            V_nodes[0] = V_dc
        else:
            V_nodes[0] = 0.0
        
        # Update inductor currents (trapezoidal)
        for k in range(n_sections):
            dV = V_nodes[k] - V_nodes[k+1] - R_sec * I_branches[k]
            I_branches[k] += (dt / L_sec) * dV
        
        # Update node voltages from capacitor currents
        for k in range(1, n_nodes):  # skip node 0 (voltage source)
            I_in = I_branches[k-1]
            I_out = I_branches[k] if k < n_sections else 0.0
            dI = I_in - I_out
            if C_nodes[k] > 0:
                V_nodes[k] += (dt / C_nodes[k]) * dI
        
        V_send_arr[idx] = V_nodes[0]
        V_recv_arr[idx] = V_nodes[-1]
        I_send_arr[idx] = I_branches[0] if n_sections > 0 else 0
        I_recv_arr[idx] = 0.0  # open end
    
    return V_send_arr, V_recv_arr, I_send_arr, I_recv_arr


def simulate_with_losses_sc(t_array, V_dc, R_t, L_t, C_t, t_sw, n_sections=1):
    """Simulate line with losses, short-circuited end."""
    dt = t_array[1] - t_array[0]
    
    R_sec = R_t / n_sections
    L_sec = L_t / n_sections
    C_sec = C_t / n_sections
    
    n_nodes = n_sections + 1
    V_nodes = np.zeros(n_nodes)
    I_branches = np.zeros(n_sections)
    
    C_nodes = np.zeros(n_nodes)
    for k in range(n_nodes):
        if k == 0 or k == n_nodes - 1:
            C_nodes[k] = C_sec / 2
        else:
            C_nodes[k] = C_sec
    
    V_send_arr = np.zeros(len(t_array))
    V_recv_arr = np.zeros(len(t_array))
    I_send_arr = np.zeros(len(t_array))
    I_recv_arr = np.zeros(len(t_array))
    
    for idx in range(len(t_array)):
        t = t_array[idx]
        
        if t >= t_sw:
            V_nodes[0] = V_dc
        else:
            V_nodes[0] = 0.0
        
        # Short circuit at last node
        V_nodes[-1] = 0.0
        
        # Update inductor currents
        for k in range(n_sections):
            dV = V_nodes[k] - V_nodes[k+1] - R_sec * I_branches[k]
            I_branches[k] += (dt / L_sec) * dV
        
        # Update node voltages (skip source node 0 and SC node n)
        for k in range(1, n_nodes - 1):
            I_in = I_branches[k-1]
            I_out = I_branches[k] if k < n_sections else 0.0
            dI = I_in - I_out
            if C_nodes[k] > 0:
                V_nodes[k] += (dt / C_nodes[k]) * dI
        
        V_send_arr[idx] = V_nodes[0]
        V_recv_arr[idx] = V_nodes[-1]
        I_send_arr[idx] = I_branches[0]
        I_recv_arr[idx] = I_branches[-1]
    
    return V_send_arr, V_recv_arr, I_send_arr, I_recv_arr


# ============================================================
# FIGURE GENERATION
# ============================================================

print("Generating figures for TP04...")

# --- Figure 1: Transmission line distributed parameter model ---
def create_fig_line_model():
    fig, ax = plt.subplots(1, 1, figsize=(13, 4))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle de la ligne à paramètres répartis (Figure 4.1)", fontsize=13, fontweight='bold')
    
    # Source
    circle = plt.Circle((1.0, 3.5), 0.35, fill=False, edgecolor='blue', lw=2)
    ax.add_patch(circle)
    ax.text(1.0, 3.5, 'E', fontsize=11, ha='center', va='center', color='blue', fontweight='bold')
    ax.text(1.0, 2.5, 'Source DC\nt=10ms', fontsize=8, ha='center', color='blue')
    
    # Switch
    ax.plot([1.35, 2.0], [3.5, 3.5], 'k-', lw=2)
    ax.plot([2.0, 2.6], [3.5, 3.8], 'k-', lw=2)
    ax.plot(2.0, 3.5, 'ko', markersize=4)
    ax.plot(2.6, 3.5, 'ko', markersize=4)
    ax.text(2.3, 4.1, 'S', fontsize=10, ha='center', fontweight='bold')
    
    # Line model (distributed parameters)
    ax.plot([2.6, 3.5], [3.5, 3.5], 'k-', lw=2)
    
    # R per unit length
    rect_r = FancyBboxPatch((3.5, 3.2), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_r)
    ax.text(4.1, 3.5, 'R·dx', fontsize=9, ha='center', va='center', fontweight='bold')
    
    # L per unit length
    ax.plot([4.7, 5.0], [3.5, 3.5], 'k-', lw=2)
    rect_l = FancyBboxPatch((5.0, 3.2), 1.2, 0.6, boxstyle="round,pad=0.03", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_l)
    ax.text(5.6, 3.5, 'L·dx', fontsize=9, ha='center', va='center', fontweight='bold')
    
    ax.plot([6.2, 7.0], [3.5, 3.5], 'k-', lw=2)
    
    # C per unit length (shunt)
    ax.plot([7.0, 7.0], [3.5, 2.5], 'k-', lw=1.5)
    ax.plot([6.7, 7.3], [2.5, 2.5], 'k-', lw=2)
    ax.plot([6.7, 7.3], [2.3, 2.3], 'k-', lw=2)
    ax.plot([7.0, 7.0], [2.3, 1.5], 'k-', lw=1.5)
    ax.text(7.5, 2.4, 'C·dx', fontsize=9, ha='left', va='center', fontweight='bold', color='green')
    
    # G per unit length (shunt conductance - often neglected)
    ax.plot([8.5, 8.5], [3.5, 2.5], 'k-', lw=1.5)
    rect_g = FancyBboxPatch((8.2, 2.0), 0.6, 0.5, boxstyle="round,pad=0.02", facecolor='#90EE90', edgecolor='black', lw=1)
    ax.add_patch(rect_g)
    ax.text(8.5, 2.25, 'G·dx', fontsize=7, ha='center', va='center')
    ax.plot([8.5, 8.5], [2.0, 1.5], 'k-', lw=1.5)
    
    # Continue line
    ax.plot([7.0, 9.5], [3.5, 3.5], 'k-', lw=2)
    ax.text(9.8, 3.5, '...', fontsize=16, ha='center', va='center')
    ax.plot([10.1, 11.5], [3.5, 3.5], 'k-', lw=2)
    
    # Receiving end
    ax.plot([11.5, 11.5], [3.5, 2.5], 'k-', lw=2)
    ax.text(12.0, 3.5, 'Extrémité\n(ouverte ou CC)', fontsize=8, ha='left', va='center', color='red')
    
    # Ground
    ax.plot([1.0, 11.5], [1.5, 1.5], 'k-', lw=2)
    ax.plot([1.0, 1.0], [1.5, 3.15], 'k-', lw=2)
    
    # Labels
    ax.text(5.5, 4.5, 'Ligne de transmission : l = 100 km', fontsize=11, ha='center', fontweight='bold', color='darkblue')
    ax.text(5.5, 0.8, r'R = 0,03 Ω/km,  L = 1 mH/km,  C = 11,11 nF/km', fontsize=10, ha='center', color='darkred')
    
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_line_model.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_line_model()
print("  - fig_line_model.png")


# --- Figure 2: Bewley lattice diagram ---
def create_fig_bewley():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))
    
    # Open end lattice
    ax = ax1
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 10)
    ax.invert_yaxis()
    ax.set_xlabel('Distance (x/l)', fontsize=10)
    ax.set_ylabel('Temps (t/τ)', fontsize=10)
    ax.set_title('Diagramme de Bewley\n(Extrémité ouverte)', fontsize=11, fontweight='bold')
    ax.axvline(x=0, color='blue', lw=2, label='Origine (source)')
    ax.axvline(x=1, color='red', lw=2, label='Extrémité (ouverte)')
    
    # Draw traveling waves
    times_s = [0, 2, 4, 6, 8]
    for i, ts in enumerate(times_s):
        ax.plot([0, 1], [ts, ts+1], 'b-', lw=1.5)
        coeff = (-1)**i
        ax.text(0.5, ts+0.3, f'V = {coeff:+d}·E', fontsize=8, color='blue', ha='center',
                rotation=-45)
    
    times_r = [1, 3, 5, 7, 9]
    for i, tr in enumerate(times_r):
        ax.plot([1, 0], [tr, tr+1], 'r-', lw=1.5)
        coeff = (-1)**i
        ax.text(0.5, tr+0.7, f'V = {coeff:+d}·E', fontsize=8, color='red', ha='center',
                rotation=45)
    
    ax.set_yticks(range(11))
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=8, loc='lower right')
    
    # Short-circuit lattice
    ax = ax2
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 10)
    ax.invert_yaxis()
    ax.set_xlabel('Distance (x/l)', fontsize=10)
    ax.set_ylabel('Temps (t/τ)', fontsize=10)
    ax.set_title('Diagramme de Bewley\n(Extrémité en court-circuit)', fontsize=11, fontweight='bold')
    ax.axvline(x=0, color='blue', lw=2, label='Origine (source)')
    ax.axvline(x=1, color='red', lw=2, label='Extrémité (CC)')
    
    for i, ts in enumerate(times_s):
        ax.plot([0, 1], [ts, ts+1], 'b-', lw=1.5)
        ax.text(0.5, ts+0.3, f'V = +E', fontsize=8, color='blue', ha='center', rotation=-45)
    
    for i, tr in enumerate(times_r):
        ax.plot([1, 0], [tr, tr+1], 'r-', lw=1.5)
        ax.text(0.5, tr+0.7, f'V = -E', fontsize=8, color='red', ha='center', rotation=45)
    
    ax.set_yticks(range(11))
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=8, loc='lower right')
    
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_bewley.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_bewley()
print("  - fig_bewley.png")


# --- Figure 3: Part 1 - Lossless line, open end ---
def create_fig_part1_open():
    dt = 1e-6
    t_max = 0.015
    t_array = np.arange(0, t_max, dt)
    
    V_send, V_recv, I_send, I_recv = simulate_distributed_param_lossless_open(
        t_array, V_source, Zc, tau, t_switch)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_send, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'origine (V_envoi)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, V_recv, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité (V_recv) - Ouverte", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_send*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine (I_envoi)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_recv*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité (I_recv) - Ouvert = 0", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    fig.suptitle("Partie 1 : Ligne sans pertes, extrémité ouverte, source DC", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part1_open.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part1_open()
print("  - fig_part1_open.png")


# --- Figure 4: Part 1 - Lossless line, short-circuited end ---
def create_fig_part1_sc():
    dt = 1e-6
    t_max = 0.015
    t_array = np.arange(0, t_max, dt)
    
    V_send, V_recv, I_send, I_recv = simulate_distributed_param_lossless_sc(
        t_array, V_source, Zc, tau, t_switch)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_send, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'origine (V_envoi)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, V_recv, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité (V_recv) - CC = 0", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_send*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine (I_envoi) - Croissant", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_recv*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité (I_recv) - CC", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 15])
    
    fig.suptitle("Partie 1 : Ligne sans pertes, extrémité en court-circuit, source DC", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part1_sc.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part1_sc()
print("  - fig_part1_sc.png")


# --- Figure 5: Part 2 - Line with losses ---
def create_fig_part2():
    dt = 5e-7
    t_max = 0.015
    t_array = np.arange(0, t_max, dt)
    
    # Open end with losses (distributed params approximated with many sections)
    V_s_o, V_r_o, I_s_o, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=50)
    V_s_sc, V_r_sc, I_s_sc, I_r_sc = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=50)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_r_o, 'r-', lw=1.5, label='Extrémité ouverte')
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité - Ligne à vide (avec pertes)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=9)
    ax.set_xlim([9.5, 15])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, I_s_o*1000, 'b-', lw=1.5, label='Ligne ouverte')
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Ligne à vide (avec pertes)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=9)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_r_sc*1000, 'r-', lw=1.5, label='Extrémité CC')
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité - Court-circuit (avec pertes)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=9)
    ax.set_xlim([9.5, 15])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_s_sc*1000, 'b-', lw=1.5, label='Courant envoi (CC)')
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Court-circuit (avec pertes)", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=9)
    ax.set_xlim([9.5, 15])
    
    fig.suptitle("Partie 2 : Ligne avec pertes (R = 0,03 Ω/km), paramètres répartis", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part2.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part2()
print("  - fig_part2.png")


# --- Figure 6: Pi model diagram ---
def create_fig_pi_model():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    
    # Single pi section
    ax = ax1
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle en π - Une section", fontsize=12, fontweight='bold')
    
    # R and L in series
    ax.plot([1, 2], [4, 4], 'k-', lw=2)
    rect_r = FancyBboxPatch((2, 3.7), 1.0, 0.6, boxstyle="round,pad=0.03", facecolor='#FFD700', edgecolor='black', lw=1.5)
    ax.add_patch(rect_r)
    ax.text(2.5, 4.0, 'R', fontsize=10, ha='center', va='center', fontweight='bold')
    
    ax.plot([3, 3.5], [4, 4], 'k-', lw=2)
    rect_l = FancyBboxPatch((3.5, 3.7), 1.0, 0.6, boxstyle="round,pad=0.03", facecolor='#ADD8E6', edgecolor='black', lw=1.5)
    ax.add_patch(rect_l)
    ax.text(4.0, 4.0, 'L', fontsize=10, ha='center', va='center', fontweight='bold')
    
    ax.plot([4.5, 6], [4, 4], 'k-', lw=2)
    
    # C/2 at input
    ax.plot([1, 1], [4, 2.8], 'k-', lw=1.5)
    ax.plot([0.7, 1.3], [2.8, 2.8], 'k-', lw=2)
    ax.plot([0.7, 1.3], [2.6, 2.6], 'k-', lw=2)
    ax.plot([1, 1], [2.6, 1.5], 'k-', lw=1.5)
    ax.text(1.6, 2.7, 'C/2', fontsize=9, ha='left', va='center', fontweight='bold', color='green')
    
    # C/2 at output
    ax.plot([6, 6], [4, 2.8], 'k-', lw=1.5)
    ax.plot([5.7, 6.3], [2.8, 2.8], 'k-', lw=2)
    ax.plot([5.7, 6.3], [2.6, 2.6], 'k-', lw=2)
    ax.plot([6, 6], [2.6, 1.5], 'k-', lw=1.5)
    ax.text(6.4, 2.7, 'C/2', fontsize=9, ha='left', va='center', fontweight='bold', color='green')
    
    # Ground
    ax.plot([1, 6], [1.5, 1.5], 'k-', lw=2)
    
    ax.text(3.5, 1.0, 'R = R\'·l, L = L\'·l, C = C\'·l', fontsize=9, ha='center', color='darkred')
    
    # Multiple sections
    ax = ax2
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Modèle en π - Plusieurs sections (n sections)", fontsize=12, fontweight='bold')
    
    for sec in range(3):
        x_off = sec * 3.5
        
        ax.plot([0.5 + x_off, 1.0 + x_off], [3.8, 3.8], 'k-', lw=2)
        rect = FancyBboxPatch((1.0 + x_off, 3.5), 0.7, 0.6, boxstyle="round,pad=0.02", facecolor='#FFD700', edgecolor='black', lw=1)
        ax.add_patch(rect)
        ax.text(1.35 + x_off, 3.8, f'R/{sec+1 if sec < 2 else "n"}' if sec > 0 else 'R/n', fontsize=7, ha='center', va='center')
        
        ax.plot([1.7 + x_off, 2.0 + x_off], [3.8, 3.8], 'k-', lw=2)
        rect2 = FancyBboxPatch((2.0 + x_off, 3.5), 0.7, 0.6, boxstyle="round,pad=0.02", facecolor='#ADD8E6', edgecolor='black', lw=1)
        ax.add_patch(rect2)
        ax.text(2.35 + x_off, 3.8, 'L/n', fontsize=7, ha='center', va='center')
        
        ax.plot([2.7 + x_off, 3.5 + x_off], [3.8, 3.8], 'k-', lw=2)
        
        # Capacitors
        ax.plot([0.5 + x_off, 0.5 + x_off], [3.8, 2.8], 'k-', lw=1)
        ax.plot([0.3 + x_off, 0.7 + x_off], [2.8, 2.8], 'k-', lw=1.5)
        ax.plot([0.3 + x_off, 0.7 + x_off], [2.6, 2.6], 'k-', lw=1.5)
        ax.plot([0.5 + x_off, 0.5 + x_off], [2.6, 2.0], 'k-', lw=1)
        
        ax.plot([3.5 + x_off, 3.5 + x_off], [3.8, 2.8], 'k-', lw=1)
        ax.plot([3.3 + x_off, 3.7 + x_off], [2.8, 2.8], 'k-', lw=1.5)
        ax.plot([3.3 + x_off, 3.7 + x_off], [2.6, 2.6], 'k-', lw=1.5)
        ax.plot([3.5 + x_off, 3.5 + x_off], [2.6, 2.0], 'k-', lw=1)
    
    ax.plot([0.5, 11.0], [2.0, 2.0], 'k-', lw=2)
    ax.text(5.5, 1.3, 'Plus le nombre de sections augmente, plus le modèle\nest précis (converge vers le modèle à paramètres répartis)',
            fontsize=8, ha='center', color='darkgreen', style='italic')
    
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_pi_model.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_pi_model()
print("  - fig_pi_model.png")


# --- Figure 7: Part 3 - Single pi section (no losses) ---
def create_fig_part3():
    dt = 1e-6
    t_max = 0.025
    t_array = np.arange(0, t_max, dt)
    
    V_s_o, V_r_o, I_s_o, _ = simulate_with_losses_open(t_array, V_source, 0, L_total, C_total, t_switch, n_sections=1)
    V_s_sc, V_r_sc, I_s_sc, I_r_sc = simulate_with_losses_sc(t_array, V_source, 0, L_total, C_total, t_switch, n_sections=1)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_r_o, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité - Ouverte (1 section π, sans pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, I_s_o*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Ouverte (1 section π, sans pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_r_sc*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité - CC (1 section π, sans pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_s_sc*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - CC (1 section π, sans pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    fig.suptitle("Partie 3 : Modèle en π - 1 section, sans pertes", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part3.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part3()
print("  - fig_part3.png")


# --- Figure 8: Part 4 - Single pi section with R, L, C ---
def create_fig_part4():
    dt = 1e-6
    t_max = 0.025
    t_array = np.arange(0, t_max, dt)
    
    V_s_o, V_r_o, I_s_o, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=1)
    V_s_sc, V_r_sc, I_s_sc, I_r_sc = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=1)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_r_o, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité - Ouverte (1 section π, avec R,L,C)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, I_s_o*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Ouverte (1 section π, avec R,L,C)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_r_sc*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité - CC (1 section π, avec R,L,C)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_s_sc*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - CC (1 section π, avec R,L,C)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 25])
    
    fig.suptitle("Partie 4 : Modèle en π - 1 section, avec pertes (R, L, C)", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part4.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part4()
print("  - fig_part4.png")


# --- Figure 9: Part 5 - 5 sections ---
def create_fig_part5():
    dt = 5e-7
    t_max = 0.020
    t_array = np.arange(0, t_max, dt)
    
    V_s_o, V_r_o, I_s_o, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=5)
    V_s_sc, V_r_sc, I_s_sc, I_r_sc = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=5)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_r_o, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité - Ouverte (5 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, I_s_o*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Ouverte (5 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_r_sc*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité - CC (5 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_s_sc*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - CC (5 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    fig.suptitle("Partie 5 : Modèle en π - 5 sections (20 km chacune), avec pertes", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part5.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part5()
print("  - fig_part5.png")


# --- Figure 10: Part 6 - 10 sections ---
def create_fig_part6():
    dt = 5e-7
    t_max = 0.020
    t_array = np.arange(0, t_max, dt)
    
    V_s_o, V_r_o, I_s_o, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=10)
    V_s_sc, V_r_sc, I_s_sc, I_r_sc = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=10)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    
    ax = axes[0, 0]
    ax.plot(t_array*1000, V_r_o, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Tension (V)', fontsize=10)
    ax.set_title("Tension à l'extrémité - Ouverte (10 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[0, 1]
    ax.plot(t_array*1000, I_s_o*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - Ouverte (10 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[1, 0]
    ax.plot(t_array*1000, I_r_sc*1000, 'r-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'extrémité - CC (10 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    ax = axes[1, 1]
    ax.plot(t_array*1000, I_s_sc*1000, 'b-', lw=1.5)
    ax.set_xlabel('Temps (ms)', fontsize=10)
    ax.set_ylabel('Courant (mA)', fontsize=10)
    ax.set_title("Courant à l'origine - CC (10 sections, avec pertes)", fontsize=10, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_xlim([9.5, 20])
    
    fig.suptitle("Partie 6 : Modèle en π - 10 sections (10 km chacune), avec pertes", fontsize=13, fontweight='bold', y=0.98)
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_part6.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_part6()
print("  - fig_part6.png")


# --- Figure 11: Comparison of all models ---
def create_fig_comparison():
    dt = 5e-7
    t_max = 0.020
    t_array = np.arange(0, t_max, dt)
    
    # Open end - receiving voltage for different models
    _, V_r_1sec, _, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=1)
    _, V_r_5sec, _, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=5)
    _, V_r_10sec, _, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=10)
    _, V_r_50sec, _, _ = simulate_with_losses_open(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=50)
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 9))
    
    ax1.plot(t_array*1000, V_r_1sec, 'g-', lw=1.5, label='1 section', alpha=0.8)
    ax1.plot(t_array*1000, V_r_5sec, 'b-', lw=1.5, label='5 sections', alpha=0.8)
    ax1.plot(t_array*1000, V_r_10sec, 'r-', lw=1.5, label='10 sections', alpha=0.8)
    ax1.plot(t_array*1000, V_r_50sec, 'k-', lw=1.5, label='50 sections (réf.)', alpha=0.8)
    ax1.set_xlabel('Temps (ms)', fontsize=11)
    ax1.set_ylabel('Tension (V)', fontsize=11)
    ax1.set_title("Comparaison : Tension à l'extrémité ouverte - Différents nombres de sections", fontsize=12, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=10)
    ax1.set_xlim([9.5, 20])
    
    # SC end - receiving current
    _, _, _, I_r_1sec = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=1)
    _, _, _, I_r_5sec = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=5)
    _, _, _, I_r_10sec = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=10)
    _, _, _, I_r_50sec = simulate_with_losses_sc(t_array, V_source, R_total, L_total, C_total, t_switch, n_sections=50)
    
    ax2.plot(t_array*1000, I_r_1sec*1000, 'g-', lw=1.5, label='1 section', alpha=0.8)
    ax2.plot(t_array*1000, I_r_5sec*1000, 'b-', lw=1.5, label='5 sections', alpha=0.8)
    ax2.plot(t_array*1000, I_r_10sec*1000, 'r-', lw=1.5, label='10 sections', alpha=0.8)
    ax2.plot(t_array*1000, I_r_50sec*1000, 'k-', lw=1.5, label='50 sections (réf.)', alpha=0.8)
    ax2.set_xlabel('Temps (ms)', fontsize=11)
    ax2.set_ylabel('Courant (mA)', fontsize=11)
    ax2.set_title("Comparaison : Courant à l'extrémité en CC - Différents nombres de sections", fontsize=12, fontweight='bold')
    ax2.grid(True, alpha=0.3)
    ax2.legend(fontsize=10)
    ax2.set_xlim([9.5, 20])
    
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_comparison.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_comparison()
print("  - fig_comparison.png")


# --- Figure 12: ATPDraw implementation diagram ---
def create_fig_atpdraw():
    fig, ax = plt.subplots(1, 1, figsize=(13, 5))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Implémentation dans ATPDraw - Ligne à paramètres répartis", fontsize=13, fontweight='bold')
    
    # DC Source
    circle = plt.Circle((1.5, 4.0), 0.4, fill=False, edgecolor='blue', lw=2)
    ax.add_patch(circle)
    ax.text(1.5, 4.0, 'DC', fontsize=9, ha='center', va='center', color='blue', fontweight='bold')
    ax.text(1.5, 3.2, 'Source\n1V DC', fontsize=8, ha='center', color='blue')
    
    # Switch (time-controlled)
    ax.plot([1.9, 2.8], [4.0, 4.0], 'k-', lw=2)
    ax.plot([2.8, 3.4], [4.0, 4.3], 'k-', lw=2)
    ax.plot(2.8, 4.0, 'ko', markersize=4)
    ax.text(3.1, 4.6, 't_close=10ms', fontsize=8, ha='center', color='purple')
    
    # Probe V1
    ax.plot([3.4, 4.0], [4.0, 4.0], 'k-', lw=2)
    ax.plot(3.7, 4.0, 'rv', markersize=8)
    ax.text(3.7, 4.4, 'V₁', fontsize=9, ha='center', color='red', fontweight='bold')
    
    # Line model (distributed parameters)
    rect_line = FancyBboxPatch((4.0, 3.3), 4.0, 1.4, boxstyle="round,pad=0.1", facecolor='#E8F4FD', edgecolor='blue', lw=2)
    ax.add_patch(rect_line)
    ax.text(6.0, 4.3, 'LIGNE', fontsize=11, ha='center', va='center', fontweight='bold', color='darkblue')
    ax.text(6.0, 3.8, 'Paramètres répartis', fontsize=9, ha='center', va='center', color='darkblue')
    ax.text(6.0, 3.5, 'l=100km, Zc=300Ω, τ=0.333ms', fontsize=7, ha='center', va='center', color='gray')
    
    # Probe V2
    ax.plot([8.0, 9.0], [4.0, 4.0], 'k-', lw=2)
    ax.plot(8.5, 4.0, 'rv', markersize=8)
    ax.text(8.5, 4.4, 'V₂', fontsize=9, ha='center', color='red', fontweight='bold')
    
    # Open end or short circuit
    ax.plot([9.0, 10.0], [4.0, 4.0], 'k-', lw=2)
    
    # Open end symbol
    ax.plot(10.0, 4.0, 'ko', markersize=6, markerfacecolor='white')
    ax.text(10.5, 4.3, 'Ouvert', fontsize=9, ha='left', color='red')
    ax.text(10.5, 3.7, '(ou CC)', fontsize=9, ha='left', color='red')
    
    # Ground connections
    ax.plot([1.5, 1.5], [3.6, 2.0], 'k-', lw=1.5)
    ax.plot([1.2, 1.8], [2.0, 2.0], 'k-', lw=2)
    ax.plot([1.3, 1.7], [1.8, 1.8], 'k-', lw=1.5)
    ax.plot([1.4, 1.6], [1.6, 1.6], 'k-', lw=1)
    
    ax.plot([10.0, 10.0], [4.0, 2.0], 'k--', lw=1.5, alpha=0.5)
    ax.plot([9.7, 10.3], [2.0, 2.0], 'k-', lw=2)
    ax.plot([9.8, 10.2], [1.8, 1.8], 'k-', lw=1.5)
    ax.text(10.0, 1.3, '(pour CC)', fontsize=7, ha='center', color='gray')
    
    # Current probes
    ax.text(3.0, 5.0, 'I₁ →', fontsize=9, ha='center', color='green', fontweight='bold')
    ax.text(8.5, 5.0, '→ I₂', fontsize=9, ha='center', color='green', fontweight='bold')
    
    # Simulation parameters box
    rect_info = FancyBboxPatch((0.5, 0.3), 5.0, 1.2, boxstyle="round,pad=0.1", facecolor='#FFFACD', edgecolor='orange', lw=1.5)
    ax.add_patch(rect_info)
    ax.text(3.0, 1.1, 'Paramètres de simulation :', fontsize=9, ha='center', fontweight='bold', color='orange')
    ax.text(3.0, 0.6, 'Δt = τ/10 ≈ 33 μs, Tmax = 15 ms', fontsize=8, ha='center')
    
    plt.tight_layout()
    plt.savefig(f"{IMG_DIR}/fig_atpdraw.png", dpi=150, bbox_inches='tight')
    plt.close()

create_fig_atpdraw()
print("  - fig_atpdraw.png")

print("All figures generated.\n")


# ============================================================
# GENERATE WORD DOCUMENT
# ============================================================

print("Generating Word document...")

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

# ---- Title Page ----
for _ in range(3):
    doc.add_paragraph('')

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('TP N°04')
run.font.size = Pt(28)
run.bold = True
run.font.color.rgb = RGBColor(0, 51, 102)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('ANALYSE TRANSITOIRE DES LIGNES DE TRANSPORT')
run.font.size = Pt(16)
run.bold = True
run.font.color.rgb = RGBColor(68, 114, 196)

doc.add_paragraph('')
subtitle2 = doc.add_paragraph()
subtitle2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle2.add_run("Modélisation dans l'EMTP - Paramètres répartis et modèle en π")
run.font.size = Pt(13)
run.italic = True
run.font.color.rgb = RGBColor(100, 100, 100)

doc.add_paragraph('')
info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = info.add_run("Département d'Électrotechnique")
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
    ('', 'Données de la ligne et paramètres calculés'),
    ('Partie 1', 'Ligne sans pertes – Modèle à paramètres répartis'),
    ('Partie 2', 'Ligne avec pertes – Modèle à paramètres répartis'),
    ('Partie 3', 'Modèle en π – 1 section sans pertes'),
    ('Partie 4', 'Modèle en π – 1 section avec R, L, C'),
    ('Partie 5', 'Modèle en π – 5 sections (20 km)'),
    ('Partie 6', 'Modèle en π – 10 sections (10 km)'),
]
for num, title_text in toc_items:
    p = doc.add_paragraph()
    run = p.add_run(f'{num}  {title_text}' if num else title_text)
    run.font.size = Pt(11)

doc.add_page_break()

# ============================================================
# PARAMETERS AND CALCULATIONS
# ============================================================

doc.add_heading('Données de la ligne et paramètres calculés', level=1)

doc.add_heading('Paramètres de la ligne', level=2)

headers = ['Paramètre', 'Symbole', 'Valeur']
rows = [
    ['Longueur de la ligne', 'l', '100 km'],
    ['Résistance linéique', "R'", '0,03 Ω/km'],
    ['Inductance linéique', "L'", '1 mH/km'],
    ['Capacité linéique', "C'", '11,11 nF/km'],
    ['Résistance totale', 'R', '3 Ω'],
    ['Inductance totale', 'L', '0,1 H = 100 mH'],
    ['Capacité totale', 'C', '1,111 μF'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Paramètres caractéristiques calculés', level=2)

p = doc.add_paragraph()
run = p.add_run("Impédance caractéristique :")
run.bold = True

doc.add_paragraph(f"   Zc = √(L'/C') = √({L_per_km*1e3:.1f}×10⁻³ / {C_per_km*1e9:.2f}×10⁻⁹) = {Zc:.1f} Ω")

doc.add_paragraph('')
p = doc.add_paragraph()
run = p.add_run("Vitesse de propagation :")
run.bold = True

doc.add_paragraph(f"   v = 1/√(L'·C') = 1/√({L_per_km*1e3:.1f}×10⁻³ × {C_per_km*1e9:.2f}×10⁻⁹) = {v_prop/1e3:.0f}×10³ km/s ≈ {v_prop:.0f} km/s")
doc.add_paragraph(f"   Soit v ≈ {v_prop/3e5*100:.0f}% de la vitesse de la lumière")

doc.add_paragraph('')
p = doc.add_paragraph()
run = p.add_run("Temps de propagation (temps de transit) :")
run.bold = True

doc.add_paragraph(f"   τ = l/v = {l_line}/{v_prop:.0f} = {tau*1e3:.4f} ms ≈ {tau*1e6:.0f} μs")

doc.add_paragraph('')
p = doc.add_paragraph()
run = p.add_run("Choix du pas de temps :")
run.bold = True

dt_chosen = tau / 10
doc.add_paragraph(f"   Δt = τ/10 = {tau*1e3:.4f}/10 = {dt_chosen*1e6:.1f} μs")
doc.add_paragraph("   Ce pas de temps assure une résolution suffisante pour capturer les fronts d'onde.")

doc.add_paragraph('')

headers = ['Paramètre calculé', 'Valeur']
rows = [
    ['Impédance caractéristique Zc', f'{Zc:.2f} Ω'],
    ['Vitesse de propagation v', f'{v_prop:.0f} km/s'],
    ['Temps de transit τ', f'{tau*1e6:.1f} μs'],
    ['Pas de temps Δt (recommandé)', f'{dt_chosen*1e6:.1f} μs'],
    ['Fréquence naturelle (ligne ouverte)', f'f₀ = 1/(4τ) = {1/(4*tau):.0f} Hz'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

doc.add_heading('Schéma du circuit (Figure 4.1)', level=2)
doc.add_picture(f"{IMG_DIR}/fig_line_model.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 1 : Modèle de la ligne de transmission à paramètres répartis")
run.italic = True
run.font.size = Pt(9)

doc.add_page_break()

# ============================================================
# PARTIE 1
# ============================================================

doc.add_heading('Partie n°1 : Ligne sans pertes – Modèle à paramètres répartis', level=1)

doc.add_heading('A. Implantation dans l\'EMTP', level=2)

p = doc.add_paragraph()
p.add_run("Pour implanter le circuit dans l'EMTP avec le modèle à paramètres répartis :").bold = True

steps = [
    "Ouvrir ATPDraw et créer un nouveau projet.",
    "Insérer un composant « Distributed Parameter Line » (LCC ou ligne à constantes réparties).",
    "Configurer les paramètres de la ligne :\n"
    f"   - Impédance caractéristique : Zc = {Zc:.2f} Ω\n"
    f"   - Temps de transit : τ = {tau*1e6:.1f} μs\n"
    "   - Pas de pertes (R = 0) pour la Partie 1",
    "Connecter une source de tension continue (DC) avec un interrupteur commandé à t = 10 ms.",
    "Laisser l'extrémité de la ligne ouverte (cas B) ou court-circuitée (cas C).",
    "Placer des sondes de tension (V) aux deux extrémités et des sondes de courant (I).",
    f"Configurer la simulation : Δt = {dt_chosen*1e6:.1f} μs, T_max = 15 ms.",
]
for i, s in enumerate(steps, 1):
    doc.add_paragraph(f"{i}. {s}")

doc.add_paragraph('')
doc.add_picture(f"{IMG_DIR}/fig_atpdraw.png", width=Inches(6.0))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 2 : Schéma d'implantation dans ATPDraw")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('B. Ligne à vide (extrémité ouverte)', level=2)

p = doc.add_paragraph()
p.add_run("Coefficients de réflexion :").bold = True

doc.add_paragraph("   • À la source (Zs = 0, source idéale) : Γ_source = (Zs - Zc)/(Zs + Zc) = -1")
doc.add_paragraph("   • À l'extrémité ouverte (Z_charge = ∞) : Γ_recv = (∞ - Zc)/(∞ + Zc) = +1")

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Analyse par le diagramme de Bewley :").bold = True

doc.add_picture(f"{IMG_DIR}/fig_bewley.png", width=Inches(6.0))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 3 : Diagrammes de Bewley (extrémité ouverte et court-circuitée)")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Question 1 : Tensions à l\'origine et à l\'extrémité', level=3)

p = doc.add_paragraph()
p.add_run("Tension à l'origine (V₁) :").bold = True
doc.add_paragraph(
    "La source de tension continue impose V₁ = E = 1 V (constante) dès la fermeture de l'interrupteur à t = 10 ms. "
    "La source idéale maintient cette tension quelle que soit la réflexion revenant de l'extrémité."
)

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Tension à l'extrémité (V₂) :").bold = True
doc.add_paragraph(
    "La tension à l'extrémité ouverte oscille entre 0 et 2E (= 2 V) avec une période de 4τ. "
    "Le phénomène se décompose ainsi :"
)

steps_v2 = [
    f"t = 10 ms + τ = 10 ms + {tau*1e3:.3f} ms : l'onde incidente arrive. La tension passe de 0 à 2E = 2 V (doublement à l'extrémité ouverte)",
    f"t = 10 ms + 3τ : l'onde réfléchie (−E) revient de la source et arrive à l'extrémité. V₂ passe de 2E à 0",
    "t = 10 ms + 5τ : une nouvelle onde +E arrive, V₂ remonte à 2E",
    "Le processus se répète indéfiniment (pas d'amortissement en l'absence de pertes)",
]
for s in steps_v2:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Question 2 : Courants à l\'origine et à l\'extrémité', level=3)

p = doc.add_paragraph()
p.add_run("Courant à l'extrémité (I₂) :").bold = True
doc.add_paragraph("   I₂ = 0 en tout temps (extrémité ouverte, pas de courant possible)")

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Courant à l'origine (I₁) :").bold = True
doc.add_paragraph(
    f"Le courant à l'origine oscille entre 0 et E/Zc = 1/{Zc:.0f} = {1/Zc*1000:.2f} mA :"
)

steps_i1 = [
    f"t = 10 ms (fermeture) : I₁ = E/Zc = {1/Zc*1000:.2f} mA (courant initial déterminé par l'impédance caractéristique)",
    f"t = 10 ms + 2τ : l'onde réfléchie (+E en tension) revient. I₁ tombe à 0",
    f"t = 10 ms + 4τ : une nouvelle onde part, I₁ remonte à E/Zc = {1/Zc*1000:.2f} mA",
    "Le courant oscille entre 0 et E/Zc avec une période 4τ",
]
for s in steps_i1:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/fig_part1_open.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 4 : Résultats de simulation - Ligne sans pertes, extrémité ouverte")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('C. Ligne court-circuitée à l\'extrémité', level=2)

p = doc.add_paragraph()
p.add_run("Coefficients de réflexion :").bold = True
doc.add_paragraph("   • À la source : Γ_source = -1 (inchangé)")
doc.add_paragraph("   • À l'extrémité en CC (Z_charge = 0) : Γ_recv = (0 - Zc)/(0 + Zc) = -1")

doc.add_paragraph('')

doc.add_heading('Question 3 : Tensions et courants', level=3)

p = doc.add_paragraph()
p.add_run("Tension à l'extrémité (V₂) :").bold = True
doc.add_paragraph("   V₂ = 0 en tout temps (imposé par le court-circuit)")

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Tension à l'origine (V₁) :").bold = True
doc.add_paragraph("   V₁ = E = 1 V (imposé par la source idéale)")

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Courant à l'origine (I₁) :").bold = True
doc.add_paragraph(
    "Le courant croît de manière illimitée par paliers de 2E/Zc toutes les 2τ :"
)
steps_isc = [
    f"t = 10 ms : I₁ = E/Zc = {1/Zc*1000:.2f} mA",
    f"t = 10 ms + 2τ : I₁ = E/Zc + 2E/Zc = 3E/Zc = {3/Zc*1000:.2f} mA",
    f"t = 10 ms + 4τ : I₁ = 5E/Zc = {5/Zc*1000:.2f} mA",
    "En général : I₁(t) = (2n+1)·E/Zc après n allers-retours",
    "Le courant croît indéfiniment (pas de pertes pour le limiter)",
]
for s in steps_isc:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run("Courant à l'extrémité (I₂) :").bold = True
doc.add_paragraph(
    "Le courant à l'extrémité court-circuitée croît également par paliers de 2E/Zc :"
)
steps_i2sc = [
    f"t = 10 ms + τ : I₂ = 2E/Zc = {2/Zc*1000:.2f} mA",
    f"t = 10 ms + 3τ : I₂ = 4E/Zc = {4/Zc*1000:.2f} mA",
    "En général : I₂(t) = 2n·E/Zc après n arrivées d'onde",
]
for s in steps_i2sc:
    doc.add_paragraph(s, style='List Bullet')

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/fig_part1_sc.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 5 : Résultats de simulation - Ligne sans pertes, extrémité court-circuitée")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Question 4 : Interprétation et comparaison', level=3)

comparison_1 = [
    "Cas ouvert : La tension à l'extrémité oscille entre 0 et 2E (surtension de 100%). Le courant à l'origine oscille entre 0 et E/Zc. L'énergie est stockée alternativement dans le champ électrique (condensateurs distribués) et le champ magnétique (inductances distribuées).",
    "Cas court-circuité : La tension à l'extrémité est nulle. Le courant augmente indéfiniment par paliers, ce qui est physiquement cohérent car l'énergie fournie par la source n'est pas dissipée (pas de résistance) et s'accumule dans le champ magnétique.",
    "La période des oscillations est la même dans les deux cas : T = 4τ (pour les tensions en circuit ouvert et les courants en CC).",
    "En l'absence de pertes, les phénomènes transitoires ne s'amortissent jamais : les oscillations persistent indéfiniment (circuit ouvert) ou le courant croît sans limite (court-circuit).",
]
for c in comparison_1:
    doc.add_paragraph(c, style='List Bullet')

doc.add_page_break()

# ============================================================
# PARTIE 2
# ============================================================

doc.add_heading('Partie n°2 : Ligne avec pertes', level=1)

doc.add_heading('Données supplémentaires', level=2)
doc.add_paragraph(f"   Résistance de la ligne : R = R' × l = {R_per_km} × {l_line} = {R_total} Ω")
doc.add_paragraph(f"   Facteur d'atténuation par transit : α = R/(2·Zc) = {R_total}/(2×{Zc:.0f}) = {R_total/(2*Zc):.4f}")
doc.add_paragraph(f"   Atténuation par transit : e^(-αl) = e^(-{R_total/(2*Zc):.4f}) ≈ {np.exp(-R_total/(2*Zc)):.4f}")

doc.add_paragraph('')

doc.add_heading('1. Résultats avec la résistance', level=2)

p = doc.add_paragraph()
p.add_run("Effet de la résistance :").bold = True

effects_r = [
    "L'atténuation des ondes voyageuses : chaque fois qu'une onde traverse la ligne, son amplitude est réduite par le facteur d'atténuation.",
    "Les oscillations (cas ouvert) ou la croissance du courant (cas CC) sont amorties exponentiellement.",
    "La tension à l'extrémité ouverte converge vers E (valeur en régime permanent) au lieu d'osciller indéfiniment entre 0 et 2E.",
    "Le courant de court-circuit converge vers E/R = 1/3 ≈ 0,333 A au lieu de croître indéfiniment.",
    "Le temps de transit reste le même (la résistance n'affecte pas la vitesse de propagation dans ce modèle simplifié).",
]
for e in effects_r:
    doc.add_paragraph(e, style='List Bullet')

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/fig_part2.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 6 : Résultats - Ligne avec pertes, paramètres répartis")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('2. Comparaison avec la Partie 1', level=2)

headers = ['Critère', 'Partie 1 (sans pertes)', 'Partie 2 (avec pertes)']
rows = [
    ['V₂ (ouvert) - régime transitoire', 'Oscille entre 0 et 2E', 'Oscille avec amortissement'],
    ['V₂ (ouvert) - régime permanent', 'Pas de convergence', 'Converge vers E'],
    ['I₁ (CC) - transitoire', 'Croissance illimitée', 'Croissance amortie'],
    ['I₁ (CC) - régime permanent', '→ ∞', f'E/R = {V_source/R_total:.3f} A'],
    ['Amortissement', 'Aucun', f'Constante de temps ≈ L/R = {L_total/R_total*1000:.1f} ms'],
    ['Surtension max (ouvert)', '2E = 2 V', '< 2E (atténuée)'],
]
add_table_with_style(doc, headers, rows)

doc.add_page_break()

# ============================================================
# PARTIE 3
# ============================================================

doc.add_heading('Partie n°3 : Modèle en π – 1 section, sans pertes', level=1)

doc.add_heading('A. Description du modèle', level=2)

p = doc.add_paragraph(
    "Le modèle en π représente la ligne par un circuit à constantes localisées. "
    "Une seule section π comprend :"
)

items_pi = [
    f"Une inductance série : L = L'×l = {L_total*1000:.0f} mH",
    f"Deux capacités shunt de C/2 chacune : C/2 = {C_total*1e6/2:.4f} μF à chaque extrémité",
    "Pas de résistance (ligne sans pertes dans cette partie)",
]
for item in items_pi:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('')
doc.add_picture(f"{IMG_DIR}/fig_pi_model.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 7 : Modèle en π – Section unique et sections multiples")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('B et C. Résultats de simulation', level=2)

p = doc.add_paragraph()
p.add_run("Fréquence de résonance du modèle à 1 section :").bold = True
f_res = 1 / (2 * np.pi * np.sqrt(L_total * C_total/2))
doc.add_paragraph(f"   f₀ = 1/(2π√(L·C/2)) = 1/(2π√({L_total:.3f}×{C_total*1e6/2:.4f}×10⁻⁶)) = {f_res:.1f} Hz")
doc.add_paragraph(f"   Période d'oscillation : T₀ = 1/f₀ = {1/f_res*1000:.2f} ms")

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/fig_part3.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 8 : Résultats - Modèle en π, 1 section, sans pertes")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Interprétation et comparaison avec la Partie 1', level=2)

comp_part3 = [
    "Le modèle à 1 section π produit des oscillations sinusoïdales pures à une seule fréquence de résonance, "
    "contrairement au modèle à paramètres répartis qui produit des fronts d'onde rectangulaires.",

    "La tension à l'extrémité ouverte oscille sinusoïdalement entre 0 et 2E (même amplitude maximale), "
    "mais avec une forme d'onde différente (sinusoïdale vs. rectangulaire).",

    "Le modèle en π ne capture pas les temps de transit finis ni les réflexions multiples. "
    "Il ne représente qu'une seule fréquence de résonance au lieu du spectre complet.",

    "La précision du modèle en π est limitée aux phénomènes basse fréquence. "
    "Pour les phénomènes rapides (fronts raides), il est inadapté avec une seule section.",
]
for c in comp_part3:
    doc.add_paragraph(c, style='List Bullet')

doc.add_page_break()

# ============================================================
# PARTIE 4
# ============================================================

doc.add_heading('Partie n°4 : Modèle en π – 1 section avec R, L, C', level=1)

doc.add_heading('A. Description du modèle', level=2)

p = doc.add_paragraph(
    "Même structure que la Partie 3, mais avec ajout de la résistance de la ligne :"
)

items_pi4 = [
    f"Résistance série : R = {R_total} Ω",
    f"Inductance série : L = {L_total*1000:.0f} mH",
    f"Capacités shunt : C/2 = {C_total*1e6/2:.4f} μF à chaque extrémité",
]
for item in items_pi4:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('B et C. Résultats de simulation', level=2)

p = doc.add_paragraph()
p.add_run("Facteur d'amortissement :").bold = True
zeta = R_total / (2 * np.sqrt(L_total / (C_total/2)))
doc.add_paragraph(f"   ζ = R/(2√(L/(C/2))) = {R_total}/(2√({L_total:.3f}/{C_total*1e6/2:.4f}×10⁻⁶)) = {zeta:.4f}")
if zeta < 1:
    doc.add_paragraph(f"   ζ < 1 → système sous-amorti : oscillations amorties")
else:
    doc.add_paragraph(f"   ζ ≥ 1 → système sur-amorti ou critique")

doc.add_paragraph('')

doc.add_picture(f"{IMG_DIR}/fig_part4.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 9 : Résultats - Modèle en π, 1 section, avec R, L, C")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Comparaison avec la Partie 3', level=2)

comp_part4 = [
    "L'ajout de la résistance introduit un amortissement des oscillations.",
    "La tension à l'extrémité ouverte converge vers E (régime permanent) au lieu d'osciller indéfiniment.",
    "Le courant de court-circuit converge vers I = E/R = 1/3 A au lieu de croître.",
    f"La constante de temps d'amortissement est τ_amort = 2L/R = 2×{L_total:.3f}/{R_total} = {2*L_total/R_total*1000:.1f} ms.",
    "La fréquence d'oscillation amortie est légèrement inférieure à la fréquence naturelle : "
    f"ω_d = ω₀√(1-ζ²) ≈ ω₀ (car ζ = {zeta:.4f} ≪ 1).",
]
for c in comp_part4:
    doc.add_paragraph(c, style='List Bullet')

doc.add_page_break()

# ============================================================
# PARTIE 5
# ============================================================

doc.add_heading('Partie n°5 : Modèle en π – 5 sections (20 km chacune)', level=1)

doc.add_heading('A. Description du modèle', level=2)

p = doc.add_paragraph(
    "La ligne de 100 km est divisée en 5 sections de 20 km chacune. "
    "Chaque section est représentée par un circuit π élémentaire avec :"
)

items_pi5 = [
    f"R par section : R/5 = {R_total/5:.2f} Ω",
    f"L par section : L/5 = {L_total/5*1000:.0f} mH",
    f"C par section : C/5 = {C_total/5*1e9:.2f} nF",
]
for item in items_pi5:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('B et C. Résultats de simulation', level=2)

doc.add_picture(f"{IMG_DIR}/fig_part5.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 10 : Résultats - Modèle en π, 5 sections (20 km)")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Comparaison avec les Parties 3 et 4', level=2)

comp_part5 = [
    "Avec 5 sections, le modèle commence à reproduire le comportement du modèle à paramètres répartis.",
    "Les fronts d'onde ne sont plus purement sinusoïdaux mais présentent des marches/escaliers "
    "qui se rapprochent des fronts raides du modèle distribué.",
    "Le délai de propagation commence à apparaître : on observe un retard entre l'application "
    f"de la tension et l'arrivée de l'onde à l'extrémité, qui se rapproche de τ = {tau*1e3:.3f} ms.",
    "Les fréquences de résonance multiples apparaissent (5 modes propres au lieu d'un seul), "
    "ce qui enrichit le contenu spectral de la réponse.",
    "La précision est nettement améliorée par rapport à 1 section, mais les hautes fréquences "
    "restent approximatives.",
    "La fréquence maximale correctement représentée est environ : "
    f"f_max ≈ v/(10×Δl) = {v_prop:.0f}/(10×20) = {v_prop/200:.0f} Hz",
]
for c in comp_part5:
    doc.add_paragraph(c, style='List Bullet')

doc.add_page_break()

# ============================================================
# PARTIE 6
# ============================================================

doc.add_heading('Partie n°6 : Modèle en π – 10 sections (10 km chacune)', level=1)

doc.add_heading('A. Description du modèle', level=2)

p = doc.add_paragraph(
    "La ligne de 100 km est divisée en 10 sections de 10 km chacune :"
)

items_pi6 = [
    f"R par section : R/10 = {R_total/10:.2f} Ω",
    f"L par section : L/10 = {L_total/10*1000:.0f} mH",
    f"C par section : C/10 = {C_total/10*1e9:.2f} nF",
]
for item in items_pi6:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('B et C. Résultats de simulation', level=2)

doc.add_picture(f"{IMG_DIR}/fig_part6.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 11 : Résultats - Modèle en π, 10 sections (10 km)")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Comparaison avec la Partie 5', level=2)

comp_part6 = [
    "Avec 10 sections, le modèle est encore plus proche du modèle à paramètres répartis.",
    "Les fronts d'onde sont plus raides et les marches d'escalier plus petites et nombreuses.",
    "Le temps de transit est mieux reproduit : le retard observé correspond de plus en plus "
    f"au τ = {tau*1e3:.3f} ms théorique.",
    "10 modes de résonance sont présents, offrant un contenu spectral plus riche.",
    "La fréquence maximale correctement représentée double : "
    f"f_max ≈ v/(10×Δl) = {v_prop:.0f}/(10×10) = {v_prop/100:.0f} Hz",
    "En pratique, au-delà de 10 sections pour 100 km, la différence avec le modèle "
    "à paramètres répartis devient négligeable pour la plupart des études transitoires.",
]
for c in comp_part6:
    doc.add_paragraph(c, style='List Bullet')

doc.add_paragraph('')

doc.add_heading('Comparaison globale – Convergence du modèle en π', level=2)

doc.add_picture(f"{IMG_DIR}/fig_comparison.png", width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Figure 12 : Comparaison de la convergence – 1, 5, 10 et 50 sections")
run.italic = True
run.font.size = Pt(9)

doc.add_paragraph('')

doc.add_heading('Synthèse et conclusion', level=2)

headers = ['Modèle', 'Avantages', 'Inconvénients']
rows = [
    ['Paramètres répartis\n(Partie 1-2)', 'Exact, capture tous les phénomènes\nde propagation', 'Disponible uniquement dans EMTP\nNe peut pas modéliser facilement\ndes lignes non uniformes'],
    ['π - 1 section\n(Partie 3-4)', 'Simple, rapide\nBon pour basse fréquence', 'Imprécis pour transitoires rapides\nUne seule fréquence de résonance'],
    ['π - 5 sections\n(Partie 5)', 'Bon compromis\nReprésente le délai de propagation', 'Limité en hautes fréquences\nPlus de composants'],
    ['π - 10 sections\n(Partie 6)', 'Très proche du modèle exact\nBonne représentation spectrale', 'Plus complexe\nTemps de calcul accru'],
]
add_table_with_style(doc, headers, rows)

doc.add_paragraph('')

conclusions = [
    "Le modèle à paramètres répartis est la référence pour les études transitoires de lignes de transport. "
    "Il reproduit exactement les phénomènes de propagation d'ondes.",

    "Le modèle en π converge vers le modèle exact lorsque le nombre de sections augmente. "
    "La règle empirique est que la longueur d'une section doit être inférieure à λ/10, "
    "où λ est la longueur d'onde minimale d'intérêt.",

    "L'ajout de pertes (résistance) amortit les phénomènes transitoires et assure la convergence "
    "vers un régime permanent. Sans pertes, les oscillations persistent indéfiniment.",

    "Pour les études de surtensions transitoires (manœuvres, foudre), le modèle à paramètres répartis "
    "ou un modèle en π avec un grand nombre de sections est indispensable.",
]
for c in conclusions:
    p = doc.add_paragraph(c)

doc.add_paragraph('')
doc.add_paragraph('')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("— Fin du document —")
run.italic = True
run.font.color.rgb = RGBColor(128, 128, 128)

# Save
output_path = "/workspace/TP04_Reponses.docx"
doc.save(output_path)
print(f"Document saved: {output_path}")
print("Done!")
