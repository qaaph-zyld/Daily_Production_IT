# Visual Dashboard Guide

## 🖥️ Dashboard Layout Preview

```
┌─────────────────────────────────────────────────────────────────┐
│                    🏭 Adient Production Dashboard                │
│              Hourly Production Monitoring - IT Department        │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 🟢 Connected  │  Last Update: 2025-01-21 13:44:00  │  ⟳ 14:32  │
└─────────────────────────────────────────────────────────────────┘

┌──────────┬──────────┬──────────┬──────────────┐
│  2,267   │    17    │   7-8    │     537      │
│  Total   │ Active   │  Peak    │    Peak      │
│Production│ Projects │  Hour    │  Production  │
└──────────┴──────────┴──────────┴──────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 📊 Production Heatmap - Hourly Distribution                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│ Project     │ 6-7│ 7-8│ 8-9│9-10│10-11│11-12│12-13│13-14│...  │
│─────────────┼────┼────┼────┼────┼─────┼─────┼─────┼─────┼─────│
│ BJA         │  0 │200 │  0 │  0 │   0 │ 200 │   0 │   0 │ 400 │
│ BR223-SEW   │  0 │ 40 │  0 │ 80 │   0 │  40 │   0 │   0 │ 160 │
│ CDPO-ASSY   │126 │ 84 │114 │ 42 │ 114 │  42 │  42 │   0 │ 564 │
│ CDPO-SEW    │ 42 │ 82 │122 │ 84 │ 122 │ 124 │  42 │   0 │ 618 │
│ FIAT-SEW    │ 60 │ 60 │153 │ 90 │ 153 │  90 │  90 │   0 │ 696 │
│ ...         │... │... │... │... │ ... │ ... │ ... │ ... │ ... │
│                                                                  │
│ [Cells are color-coded: Light Blue → Dark Blue based on value]  │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 📈 Total Production by Project                                   │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│ MMA-ASSY    ████████████████████████████ 2,267                  │
│ MMA-SEW     ████████████████████ 1,320                          │
│ VOLVO-ASSY  ████████████████ 976                                │
│ VOLVO-SEW   ███████████████ 1,007                               │
│ FIAT-SEW    ██████████████ 696                                  │
│ CDPO-SEW    █████████████ 618                                   │
│ CDPO-ASSY   ████████████ 564                                    │
│ SCANIA      ██████████ 470                                      │
│ MAN         █████████ 624                                       │
│ BJA         ████ 400                                            │
│ ...         ...                                                 │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ ⏰ Production by Time Slot                                       │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│  600│                    ╱╲                                      │
│     │                   ╱  ╲                                     │
│  500│                  ╱    ╲                                    │
│     │                 ╱      ╲                                   │
│  400│          ╱╲    ╱        ╲                                  │
│     │         ╱  ╲  ╱          ╲                                 │
│  300│        ╱    ╲╱            ╲                                │
│     │       ╱                    ╲                               │
│  200│      ╱                      ╲                              │
│     │     ╱                        ╲                             │
│  100│    ╱                          ╲                            │
│     │   ╱                            ╲                           │
│    0└───┴───┴───┴───┴───┴───┴───┴───┴───┴───                    │
│       6-7 7-8 8-9 9-10 10-11 11-12 12-13 13-14 14-22 II         │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 📋 Detailed Production Table                                     │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│ [Same data as heatmap but in standard table format]             │
│ [Scrollable, with sticky headers]                               │
│ [Hover effects on rows]                                         │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🎨 Color Scheme Visualization

### Background Gradient
```
Top:    #0f4c75 (Deep Blue)
        ↓ Gradient ↓
Bottom: #1b262c (Navy)
```

### Card Colors
```
Header/Cards: rgba(255, 255, 255, 0.1) with blur
Text:         #bbe1fa (Light Blue)
Accents:      #3282b8 (Ocean Blue)
```

### Heatmap Colors
```
Low (0):      rgb(187, 225, 250) - Light Blue
Medium (50%): rgb(101, 150, 183) - Medium Blue
High (100%):  rgb(15, 76, 117)   - Dark Blue
```

### Chart Colors
```
Bar Chart:  Multi-color palette (17 colors)
Line Chart: #3282b8 with 0.2 opacity fill
```

---

## 📊 Component Breakdown

### 1. Header Section
```
┌─────────────────────────────────────────┐
│  🏭 Adient Production Dashboard         │
│  Hourly Production Monitoring - IT Dept │
└─────────────────────────────────────────┘
```
- **Background**: Glass-morphism effect
- **Font**: 2.5em bold, light blue
- **Shadow**: Text shadow for depth

### 2. Status Bar
```
┌──────────────────────────────────────────────────────┐
│ 🟢 Connected │ Last Update: ... │ ⟳ Next refresh: ...│
└──────────────────────────────────────────────────────┘
```
- **Indicator**: Pulsing green dot (or red if error)
- **Layout**: Flexbox, space-between
- **Updates**: Real-time countdown

### 3. Statistics Cards (Grid)
```
┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐
│  2,267   │ │    17    │ │   7-8    │ │   537    │
│  Total   │ │ Active   │ │  Peak    │ │  Peak    │
│Production│ │ Projects │ │  Hour    │ │Production│
└──────────┘ └──────────┘ └──────────┘ └──────────┘
```
- **Layout**: CSS Grid, auto-fit
- **Background**: Blue gradient
- **Font**: 2em value, 0.9em label

### 4. Heatmap Table
```
┌────────────┬─────┬─────┬─────┬─────┐
│ Project    │ 6-7 │ 7-8 │ 8-9 │ ... │
├────────────┼─────┼─────┼─────┼─────┤
│ BJA        │  0  │ 200 │  0  │ ... │
│ BR223-SEW  │  0  │  40 │  0  │ ... │
│ CDPO-ASSY  │ 126 │  84 │ 114 │ ... │
└────────────┴─────┴─────┴─────┴─────┘
```
- **Cells**: Color intensity based on value
- **Hover**: Scale transform (1.05)
- **Header**: Sticky, dark blue background

### 5. Bar Chart (Chart.js)
```
MMA-ASSY    ████████████████████████████
MMA-SEW     ████████████████████
VOLVO-ASSY  ████████████████
VOLVO-SEW   ███████████████
```
- **Type**: Horizontal bar
- **Height**: 600px
- **Interactive**: Tooltips on hover
- **Responsive**: Maintains aspect ratio

### 6. Line Chart (Chart.js)
```
     ╱╲
    ╱  ╲
   ╱    ╲
  ╱      ╲
 ╱        ╲
```
- **Type**: Line with area fill
- **Tension**: 0.4 (smooth curves)
- **Points**: 6px radius, 8px on hover
- **Fill**: Blue gradient

---

## 🖱️ Interactive Elements

### Hover Effects
- **Cards**: Subtle shadow increase
- **Table Rows**: Background color change
- **Heatmap Cells**: Scale up (1.05x)
- **Chart Points**: Radius increase + tooltip

### Click Interactions
- **Charts**: Click legend to toggle datasets
- **Tables**: Click headers to sort (future)

### Animations
- **Status Indicator**: Pulse animation (2s loop)
- **Page Load**: Fade-in effects
- **Data Update**: Smooth transitions

---

## 📱 Responsive Breakpoints

### Desktop (1920px+)
```
┌─────────────────────────────────────┐
│ [Full width, side-by-side stats]    │
│ [Large charts, optimal spacing]     │
└─────────────────────────────────────┘
```

### Tablet (768px - 1919px)
```
┌──────────────────────┐
│ [Stacked layout]     │
│ [Medium charts]      │
│ [Readable tables]    │
└──────────────────────┘
```

### Mobile (< 768px)
```
┌──────────┐
│ [Single] │
│ [Column] │
│ [Compact]│
│ [Scroll] │
└──────────┘
```

---

## 🎭 Visual States

### Loading State
```
┌─────────────────────┐
│                     │
│   Loading data...   │
│                     │
└─────────────────────┘
```

### Error State
```
┌─────────────────────────────────┐
│ ⚠️ Error Loading Data           │
│ [Error message]                 │
│ Please check database connection│
└─────────────────────────────────┘
```

### Success State
```
┌─────────────────────────────────┐
│ 🟢 Connected                    │
│ [All data displayed]            │
│ [Charts rendered]               │
└─────────────────────────────────┘
```

---

## 🌈 Accessibility Features

- **High Contrast**: Text easily readable
- **Color + Shape**: Not relying on color alone
- **Large Touch Targets**: Mobile-friendly
- **Clear Labels**: All data identified
- **Semantic HTML**: Screen reader friendly

---

## 💡 Visual Tips

1. **Full Screen**: Press F11 for immersive view
2. **Zoom**: Ctrl + Mouse Wheel to adjust size
3. **Print**: Ctrl + P for PDF export
4. **Screenshot**: Use Snipping Tool for reports
5. **Multiple Monitors**: Dedicate one screen to dashboard

---

## 🎨 Customization Examples

### Change Heatmap Colors (Green Theme)
```javascript
// In dashboard.html, modify getHeatmapColor()
const r = Math.round(200 - (200 - 50) * intensity);
const g = Math.round(255 - (255 - 150) * intensity);
const b = Math.round(200 - (200 - 50) * intensity);
```

### Adjust Card Gradient
```css
.stat-card {
    background: linear-gradient(135deg, #4CAF50 0%, #2E7D32 100%);
}
```

### Modify Chart Colors
```javascript
const colors = [
    '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
    '#9966FF', '#FF9F40', '#FF6384', '#C9CBCF'
];
```

---

**This visual guide helps you understand the dashboard layout and design before launching it!**
