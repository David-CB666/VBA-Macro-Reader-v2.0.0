#!/usr/bin/env python3
"""Hero image for VBA Macro Reader: dual-mode code diagram."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from pathlib import Path

def render(out_path):
    fig, ax = plt.subplots(figsize=(12, 5.5), dpi=150)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 5.5)
    ax.axis('off')
    
    DARK = '#1F3864'
    BLUE = '#2E75B6'
    GREEN = '#2D8B4A'
    ORANGE = '#ED7D31'
    RED = '#C00000'
    GRAY = '#808080'
    WHITE = '#FFFFFF'
    LIGHT_BLUE = '#D6E4F0'
    LIGHT_GREEN = '#E8F5E9'
    BG = '#F8F9FB'
    
    fig.patch.set_facecolor(BG)
    
    # Title
    ax.text(0.5, 5.15, 'VBA Macro Reader & Operator', fontsize=22, fontweight='bold',
           color=DARK, ha='left')
    ax.text(0.55, 4.85, 'Read · Modify · Execute VBA macros from .xlsm files', fontsize=12,
           color='#666', ha='left', style='italic')
    
    # --- Input box ---
    box_in = FancyBboxPatch((1.0, 3.5), 1.8, 0.9, boxstyle="round,pad=0.12",
                            facecolor=WHITE, edgecolor=DARK, linewidth=1.5, zorder=2)
    ax.add_patch(box_in)
    ax.text(1.9, 4.2, '.xlsm / .xlam', fontsize=12, fontweight='bold', color=DARK, ha='center')
    ax.text(1.9, 3.85, 'Macro-enabled', fontsize=8, color='#999', ha='center')
    ax.text(1.9, 3.68, 'Excel workbook', fontsize=8, color='#999', ha='center')
    
    # Arrow in -> mode selector
    arrow1 = FancyArrowPatch((2.9, 3.95), (3.7, 3.95), arrowstyle='->', color=GRAY,
                             linewidth=1.5, mutation_scale=15, zorder=1)
    ax.add_patch(arrow1)
    
    # Mode selector
    box_sel = FancyBboxPatch((3.8, 3.5), 1.8, 0.9, boxstyle="round,pad=0.12",
                             facecolor=WHITE, edgecolor=ORANGE, linewidth=1.5, zorder=2)
    ax.add_patch(box_sel)
    ax.text(4.7, 4.2, 'Mode Select', fontsize=10, fontweight='bold', color=ORANGE, ha='center')
    ax.text(4.7, 3.88, 'Auto-detect', fontsize=8, color='#555', ha='center')
    ax.text(4.7, 3.68, 'or specify', fontsize=8, color='#999', ha='center')
    
    # --- Path 1: oletools (no Excel) ---
    arrow_o = FancyArrowPatch((5.7, 4.25), (6.8, 4.65), arrowstyle='->', color=GRAY,
                              linewidth=1.2, mutation_scale=12, connectionstyle="arc3,rad=0.3", zorder=1)
    ax.add_patch(arrow_o)
    arrow_o2 = FancyArrowPatch((5.7, 3.65), (6.8, 3.25), arrowstyle='->', color=GRAY,
                               linewidth=1.2, mutation_scale=12, connectionstyle="arc3,rad=-0.3", zorder=1)
    ax.add_patch(arrow_o2)
    
    # oletools mode
    box_o = FancyBboxPatch((6.9, 4.05), 2.4, 1.2, boxstyle="round,pad=0.12",
                           facecolor=LIGHT_BLUE, edgecolor=BLUE, linewidth=1.5, zorder=2)
    ax.add_patch(box_o)
    ax.text(8.1, 4.95, 'oletools Mode', fontsize=11, fontweight='bold', color=DARK, ha='center')
    ax.text(8.1, 4.65, 'Read VBA code', fontsize=8.5, color='#555', ha='center')
    ax.text(8.1, 4.45, 'Cross-platform', fontsize=7.5, color='#999', ha='center')
    ax.text(8.1, 4.25, 'No Excel needed', fontsize=7.5, color='#999', ha='center')
    
    # Badge on oletools
    ax.text(8.1, 4.68, '', fontsize=0)  # spacer
    badge_o = FancyBboxPatch((7.6, 4.15), 1.0, 0.25, boxstyle="round,pad=0.05",
                             facecolor=BLUE, edgecolor='none', alpha=0.15, zorder=3)
    ax.add_patch(badge_o)
    ax.text(8.1, 4.27, 'READ', fontsize=7, fontweight='bold', color=BLUE, ha='center', zorder=4)
    
    # win32com mode
    box_w = FancyBboxPatch((6.9, 2.65), 2.4, 1.2, boxstyle="round,pad=0.12",
                           facecolor=LIGHT_GREEN, edgecolor=GREEN, linewidth=1.5, zorder=2)
    ax.add_patch(box_w)
    ax.text(8.1, 3.55, 'win32com Mode', fontsize=11, fontweight='bold', color=DARK, ha='center')
    ax.text(8.1, 3.25, 'Read + Modify + Execute', fontsize=8.5, color='#555', ha='center')
    ax.text(8.1, 3.05, 'Windows + Excel required', fontsize=7.5, color='#999', ha='center')
    ax.text(8.1, 2.85, 'Full VBA control', fontsize=7.5, color='#999', ha='center')
    
    # Badges on win32com
    badges_data = [('READ', BLUE), ('WRITE', ORANGE), ('EXEC', GREEN)]
    for j, (label, color) in enumerate(badges_data):
        bx = 7.15 + j * 0.65
        by = 2.75
        badge = FancyBboxPatch((bx, by), 0.55, 0.25, boxstyle="round,pad=0.04",
                               facecolor=color, edgecolor='none', alpha=0.12, zorder=3)
        ax.add_patch(badge)
        ax.text(bx + 0.275, by + 0.125, label, fontsize=6, fontweight='bold',
               color=color, ha='center', zorder=4)
    
    # --- Right side: results ---
    arrow_res1 = FancyArrowPatch((9.4, 4.65), (10.3, 3.95), arrowstyle='->', color=GRAY,
                                 linewidth=1.2, mutation_scale=12, connectionstyle="arc3,rad=-0.3", zorder=1)
    ax.add_patch(arrow_res1)
    arrow_res2 = FancyArrowPatch((9.4, 3.25), (10.3, 3.95), arrowstyle='->', color=GRAY,
                                 linewidth=1.2, mutation_scale=12, connectionstyle="arc3,rad=0.3", zorder=1)
    ax.add_patch(arrow_res2)
    
    # Output
    box_out = FancyBboxPatch((10.4, 3.5), 1.2, 0.9, boxstyle="round,pad=0.12",
                             facecolor=DARK, edgecolor=DARK, linewidth=1.5, zorder=2)
    ax.add_patch(box_out)
    ax.text(11.0, 4.15, 'Results', fontsize=9, fontweight='bold', color=WHITE, ha='center')
    ax.text(11.0, 3.9, 'Module list', fontsize=7, color='#AAC4DF', ha='center')
    ax.text(11.0, 3.75, 'Code + analysis', fontsize=7, color='#AAC4DF', ha='center')
    ax.text(11.0, 3.6, 'Output files', fontsize=7, color='#AAC4DF', ha='center')
    
    # --- Feature list bottom ---
    features = [
        ('Extract modules', 'All VBA components'),
        ('Analyze structure', 'Sub/Function/Params'),
        ('Modify code', 'Dynamic updates via COM'),
        ('Execute macros', 'Run + monitor + timeout'),
        ('Batch processing', 'Chain macros, handle results'),
        ('CI-ready tests', '7 pytest cases'),
    ]
    
    for i, (title, desc) in enumerate(features):
        x = 0.5 + i * 2.0
        ax.text(x, 1.8, '✓', fontsize=12, color=GREEN, fontweight='bold', ha='left')
        ax.text(x + 0.3, 1.8, title, fontsize=8, fontweight='bold', color=DARK, ha='left', va='center')
        ax.text(x + 0.3, 1.57, desc, fontsize=7, color='#888', ha='left', va='center')
    
    # Tagline
    ax.text(0.5, 0.6, 'pip install pywin32 oletools  →  from vba_reader import VBAReader  →  reader.run_macro("...")',
           fontsize=9, color='#aaa', style='italic', ha='left')
    
    plt.tight_layout(pad=0.5)
    fig.savefig(out_path, dpi=150, bbox_inches='tight', facecolor=BG, edgecolor='none')
    plt.close(fig)
    print(f"✓ Saved: {out_path} ({Path(out_path).stat().st_size // 1024} KB)")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "demo/demo_preview.png"
    render(out)
