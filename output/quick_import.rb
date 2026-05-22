# Floorplan Import Script - Quick Version
# Save this file somewhere, then in SketchUp Ruby Console type:
# load '/Users/vega/CascadeProjects/windsurf-project/output/quick_import.rb'

require 'sketchup.rb'
require 'extensions.rb'

module FloorplanImport
  def self.run
    model = Sketchup.active_model
    model.start_operation('Import Floorplan', true)
    
    entities = model.entities
    
    # Materials
    mats = model.materials
    wall_mat = mats.add('Wall')
    wall_mat.color = Sketchup::Color.new(240, 235, 228)
    
    floor_mat = mats.add('Floor')
    floor_mat.color = Sketchup::Color.new(74, 74, 74)
    
    # Convert meters to inches
    def self.m(meters)
      meters * 39.3701
    end
    
    # Create floors - Main Building
    floors = [
      # [x1, z1, width, depth, name]
      [0, 0, 5.53, 4.0, "Slaapkamer 1"],
      [0, 4.0, 3.66, 3.06, "Badkamer"],
      [0, 7.06, 3.66, 3.3, "Slaapkamer 2"],
      [0, 10.36, 1.15, 1.5, "Toilet"],
      [0, 11.86, 1.15, 2.22, "CV"],
      [3.66, 4.0, 1.37, 5.82, "Foyer"],
      [1.15, 10.36, 2.51, 3.72, "Entrance"],
      [3.66, 9.82, 7.68, 5.58, "Keuken"],
      [11.34, 9.82, 3.1, 5.58, "Woonkamer"],
    ]
    
    # Garage/Office
    garage_floors = [
      [18.44, 3.12, 5.26, 3.87, "Office"],
      [18.44, 6.99, 5.26, 5.29, "Garage"],
    ]
    
    all_floors = floors + garage_floors
    
    all_floors.each do |x1, z1, w, d, name|
      pts = [
        [m(x1), 0, m(z1)],
        [m(x1 + w), 0, m(z1)],
        [m(x1 + w), 0, m(z1 + d)],
        [m(x1), 0, m(z1 + d)]
      ]
      face = entities.add_face(pts)
      face.material = floor_mat if face
    end
    
    # Outer walls - Horizontal
    outer_h_walls = [
      # [x1, x2, z, height]
      [0, 5.53, 0, 2.8],      # North slaapkamer 1
      [5.03, 14.44, 4.0, 2.8], # North keuken area
      [3.66, 14.44, 15.4, 2.8], # South main
      [0, 3.66, 14.08, 2.8],    # South left
      [18.44, 23.70, 3.12, 2.8], # North garage
      [18.44, 23.70, 12.28, 2.46], # South garage
    ]
    
    # Outer walls - Vertical
    outer_v_walls = [
      # [x, z1, z2, height]
      [0, 0, 14.08, 2.8],       # West main
      [14.44, 4.0, 15.4, 2.8],  # East main
      [18.44, 3.12, 12.28, 2.46], # West garage
      [23.70, 3.12, 12.28, 2.46], # East garage
    ]
    
    thick = m(0.2)
    
    # Build horizontal walls
    outer_h_walls.each do |x1, x2, z, h|
      pts = [
        [m(x1), 0, m(z) - thick/2],
        [m(x2), 0, m(z) - thick/2],
        [m(x2), 0, m(z) + thick/2],
        [m(x1), 0, m(z) + thick/2]
      ]
      face = entities.add_face(pts)
      if face
        face.pushpull(-m(h))
        face.material = wall_mat
      end
    end
    
    # Build vertical walls
    outer_v_walls.each do |x, z1, z2, h|
      pts = [
        [m(x) - thick/2, 0, m(z1)],
        [m(x) + thick/2, 0, m(z1)],
        [m(x) + thick/2, 0, m(z2)],
        [m(x) - thick/2, 0, m(z2)]
      ]
      face = entities.add_face(pts)
      if face
        face.pushpull(-m(h))
        face.material = wall_mat
      end
    end
    
    model.commit_operation
    UI.messagebox("Floorplan imported successfully!\n\nSwitch to Top view to see the layout.")
  end
end

# Run immediately
FloorplanImport.run
