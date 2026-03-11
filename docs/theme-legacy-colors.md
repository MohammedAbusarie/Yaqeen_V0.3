# Legacy Theme: Pre–March 2026 (Blue / Dark Blue-Grey)

This document records the **previous website identity and color palette** before the March 2026 rebrand to the dark black, red, and white theme. It is kept for reference (e.g. reverting, design history, or brand audits).

**Replaced:** March 2026  
**Current theme:** Dark black, red, white (see `.project/memory/manifest.md` and `styles.css` `:root`).

---

## CSS Custom Properties (`:root`)

### Backgrounds
| Variable | Value | Description |
|----------|--------|-------------|
| `--bg` | `#0a0e1a` | Main page background (dark blue-grey) |
| `--bg-gradient-1` | `#1a1f35` | Top gradient layer |
| `--bg-gradient-2` | `#0f1420` | Bottom gradient layer |

### Surfaces (cards, panels, inputs)
| Variable | Value | Description |
|----------|--------|-------------|
| `--surface` | `#151b2e` | Card/panel base |
| `--surface-elevated` | `#1c2439` | Raised surfaces (inputs, table headers) |
| `--surface-hover` | `#222948` | Hover state for surfaces |

### Text
| Variable | Value | Description |
|----------|--------|-------------|
| `--text-primary` | `#f0f4ff` | Primary text (off-white with blue tint) |
| `--text-secondary` | `#9ca8d4` | Secondary text (muted blue-grey) |
| `--text-muted` | `#6b7599` | Muted labels, hints |

### Borders
| Variable | Value | Description |
|----------|--------|-------------|
| `--border` | `rgba(148, 163, 207, 0.12)` | Default border (blue-grey, low opacity) |
| `--border-focus` | `rgba(99, 155, 255, 0.4)` | Focus ring (blue) |

### Brand / Primary (Blue)
| Variable | Value | Description |
|----------|--------|-------------|
| `--primary` | `#639bff` | Primary actions, links, active nav |
| `--primary-hover` | `#7aaeff` | Primary hover |
| `--primary-active` | `#5088e6` | Primary active/pressed |
| `--primary-subtle` | `rgba(99, 155, 255, 0.12)` | Subtle primary tint (code, highlights) |

### Semantic Colors
| Variable | Value | Description |
|----------|--------|-------------|
| `--success` | `#4ade80` | Success state (green) |
| `--success-subtle` | `rgba(74, 222, 128, 0.12)` | Success background tint |
| `--danger` | `#f87171` | Error/destructive (red) |
| `--danger-subtle` | `rgba(248, 113, 113, 0.12)` | Error background tint |
| `--warning` | `#fbbf24` | Warning (amber) |
| `--warning-subtle` | `rgba(251, 191, 36, 0.12)` | Warning background tint |

### Shadows
| Variable | Value |
|----------|--------|
| `--shadow-sm` | `0 2px 8px rgba(0, 0, 0, 0.12)` |
| `--shadow-md` | `0 4px 16px rgba(0, 0, 0, 0.18)` |
| `--shadow-lg` | `0 8px 32px rgba(0, 0, 0, 0.28)` |
| `--shadow-xl` | `0 16px 48px rgba(0, 0, 0, 0.35)` |

### Border radius & transitions
- `--radius-sm`: 8px  
- `--radius-md`: 12px  
- `--radius-lg`: 16px  
- Transitions: 150ms / 250ms / 350ms cubic-bezier(0.4, 0, 0.2, 1)

---

## Body Background (Gradients)

```css
background:
  radial-gradient(circle at 15% 10%, rgba(99, 155, 255, 0.15), transparent 45%),
  radial-gradient(circle at 85% 15%, rgba(74, 222, 128, 0.08), transparent 50%),
  radial-gradient(circle at 50% 100%, rgba(99, 155, 255, 0.05), transparent 60%),
  linear-gradient(180deg, var(--bg-gradient-1) 0%, var(--bg) 100%);
```

- Blue glow top-left and bottom-center; green glow top-right.

---

## Card Accents

- **Card top edge (`.card::before`):**  
  `linear-gradient(90deg, transparent, rgba(99, 155, 255, 0.5), transparent)`
- **Card hover glow (`.card::after`):**  
  `radial-gradient(circle, rgba(99, 155, 255, 0.03) 0%, transparent 70%)`

---

## Other Hardcoded Blue/Green (Legacy)

- **Button hover border:** `rgba(148, 163, 207, 0.3)`
- **Primary button shadow:** `rgba(99, 155, 255, 0.3)` / `rgba(99, 155, 255, 0.4)`
- **File input button hover:** `rgba(99, 155, 255, 0.2)`
- **Status success border:** `rgba(74, 222, 128, 0.3)`
- **Table row--title hover:** `rgba(99, 155, 255, 0.18)`
- **Scrollbar thumb hover:** `rgba(148, 163, 207, 0.3)`
- **Splash screen background:** `radial-gradient(circle at 50% 50%, rgba(99, 155, 255, 0.2), transparent 70%)`
- **Splash screen title:** gradient `var(--text-primary)` → `var(--primary-hover)` → `var(--primary)`; `text-shadow: 0 0 40px rgba(99, 155, 255, 0.3)`
- **Details card (modal):** `rgba(21, 27, 46, 0.35)` / `0.55` (blue-tinted dark)

---

## Summary

- **Identity:** Dark theme with **blue** as primary/accent and blue-grey surfaces.
- **Hex reference:** Primary blue `#639bff` (RGB 99, 155, 255).
- **Contrast:** Off-white text on dark blue-grey; blue used for interactive elements and focus.

For the current theme values, see `styles.css` (`:root`) and `.project/memory/manifest.md`.
