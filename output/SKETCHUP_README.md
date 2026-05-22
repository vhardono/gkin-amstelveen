# Floorplan for SketchUp Import

## Quick Reference

### Units
- **All measurements in meters**
- Import/Model in meters scale
- Wall height: 2.8m (garage: 2.46m)
- Wall thickness: 0.2m

### Building Layout

#### MAIN BUILDING (L-Shape)
Coordinates are from top-left corner of the building:

```
Z=0 (North)
  ←───────────────── 14.44m ─────────────────→
  ┌─────────┬────────┬─────────────────────────┐
  │Slaapk. 1│Badkamer│ Slaapkamer 2 │ Toilet │CV │
  │ 5.53×4m │3.66×3m │ 3.66×3.3m   │1.15×1.5│...│
  │         │        │             │        │   │
  │ ◻️◻️    │  ⬆️    │   ◻️        │        │   │ ← West Wall
  │ windows │ high   │  window     │        │   │   (x=0)
  │         │ window │             │        │   │
  ├────┬────┼────────┴─────────────┴────────┴───┤ Z=4m
  │    │Foyer│                               │
  │    │1.37m│                               │
  │    │wide │      Keuken / Woonkamer       │
  │Ent │corri│     10.78m × 5.58m            │
  │rance│dor│                               │
  │2.5m│    │  ◻️◻️◻️        ◻️    ◻️       │
  │×3.7│    │  windows      win   win       │
  │    │    │                               │
  └────┴────┴───────────────────────────────┘ Z=15.4m (South)
       ↑
      Entrance porch
```

#### GARAGE/OFFICE (Separate Building)
- **Location**: 4m gap east of main building
- **Dimensions**: 5.26m (W) × 9.16m (D)
- **Split**: Office (north half), Garage (south half)

```
  X=18.44              X=23.70
  ←──── 5.26m ──────→
  ┌─────────────────┐ Z=3.12
  │     Office      │ 3.87m deep
  │  ◻️◻️◻️◻️ ◻️   │
  │  4m window      │
  ├─────────────────┤ Z=6.99 (divider)
  │     Garage      │ 5.29m deep
  │   [====] door   │ 4m garage door
  │      ◻️         │ side door
  └─────────────────┘ Z=12.28
```

## Import Steps

1. **Open SketchUp** → Set template to "Meters"

2. **Import CSV** (if using extension):
   - Use `floorplan-rooms.csv` for room rectangles
   - Use `floorplan-windows.csv` for window openings

3. **Manual Modeling**:
   - Draw rooms using rectangle tool with coordinates from CSV
   - Push/Pull walls to 2.8m height
   - Draw windows as rectangles and push through walls
   - Use components for repeating elements (windows, doors)

4. **Recommended Components to Create**:
   - Window FH (full height): 1m × 2.8m
   - Window high: 1m × 0.6m @ 1.8m sill
   - Front door: 1m × 2.8m
   - Garage door: 4m × 2.26m
   - Glass wall panel: 1m × 2.8m (for foyer/keuken)

## Key Details

- **Plinth**: Add 0.3m protrusion at base of all outer walls
- **Floors**: 
  - Default rooms: standard floor
  - Badkamer/Toilet: dark tile floor
  - Garage: concrete floor
- **Foyer glass wall**: 3 full-height panes (east wall x=5.03)
- **Keuken glass wall**: 3 full-height panes (north wall z=4.00)

## Export Back

When done, export from SketchUp as:
1. **3D Model**: `.obj` or `.dae` (Collada) format
2. **2D Plans**: `.svg` or `.png` screenshots
3. **Dimensions**: Update the CSV files with any changes

Send back any combination of these and I can update the web viewer!
