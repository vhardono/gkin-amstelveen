# Floorplan Import Script - Fixed Version
# Run in Ruby Console: load '/Users/vega/CascadeProjects/windsurf-project/output/floorplan_fixed.rb'

module FloorplanImport
  def self.run
    model = Sketchup.active_model
    model.start_operation('Import Floorplan', true)
    
    entities = model.entities
    
    # Create group first to isolate geometry
    group = entities.add_group
    gents = group.entities
    
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
    
    thick = m(0.2)
    
    puts "Creating floors..."
    
    # Create floors - all rooms
    floors = [
      [0, 0, 5.53, 4.0],        # Slaapkamer 1
      [0, 4.0, 3.66, 3.06],     # Badkamer
      [0, 7.06, 3.66, 3.3],     # Slaapkamer 2
      [0, 10.36, 1.15, 1.5],    # Toilet
      [0, 11.86, 1.15, 2.22],   # CV
      [3.66, 4.0, 1.37, 5.82],  # Foyer
      [1.15, 10.36, 2.51, 3.72], # Entrance
      [3.66, 9.82, 7.68, 5.58], # Keuken
      [11.34, 9.82, 3.1, 5.58], # Woonkamer
      [18.44, 3.12, 5.26, 3.87], # Office
      [18.44, 6.99, 5.26, 5.29], # Garage
    ]
    
    floors.each do |x1, z1, w, d|
      pts = [
        [m(x1), 0, m(z1)],
        [m(x1 + w), 0, m(z1)],
        [m(x1 + w), 0, m(z1 + d)],
        [m(x1), 0, m(z1 + d)]
      ]
      face = gents.add_face(pts)
      face.material = floor_mat if face && face.valid?
    end
    
    puts "Creating outer walls..."
    
    # Create walls as separate groups to avoid merging issues
    walls_group = gents.add_group
    wall_ents = walls_group.entities
    
    # Horizontal walls (facing Z)
    h_walls = [
      [0, 5.53, 0, 2.8],
      [5.03, 14.44, 4.0, 2.8],
      [3.66, 14.44, 15.4, 2.8],
      [0, 3.66, 14.08, 2.8],
      [18.44, 23.70, 3.12, 2.8],
      [18.44, 23.70, 12.28, 2.46],
    ]
    
    h_walls.each do |x1, x2, z, h|
      # Create a small group for each wall
      wg = wall_ents.add_group
      we = wg.entities
      
      pts = [
        [m(x1), 0, m(z) - thick/2],
        [m(x2), 0, m(z) - thick/2],
        [m(x2), 0, m(z) + thick/2],
        [m(x1), 0, m(z) + thick/2]
      ]
      face = we.add_face(pts)
      if face && face.valid?
        face.pushpull(-m(h))
        # Apply material to all faces in the wall
        we.grep(Sketchup::Face).each { |f| f.material = wall_mat if f.valid? }
      end
    end
    
    # Vertical walls (facing X)
    v_walls = [
      [0, 0, 14.08, 2.8],
      [14.44, 4.0, 15.4, 2.8],
      [18.44, 3.12, 12.28, 2.46],
      [23.70, 3.12, 12.28, 2.46],
    ]
    
    v_walls.each do |x, z1, z2, h|
      wg = wall_ents.add_group
      we = wg.entities
      
      pts = [
        [m(x) - thick/2, 0, m(z1)],
        [m(x) + thick/2, 0, m(z1)],
        [m(x) + thick/2, 0, m(z2)],
        [m(x) - thick/2, 0, m(z2)]
      ]
      face = we.add_face(pts)
      if face && face.valid?
        face.pushpull(-m(h))
        we.grep(Sketchup::Face).each { |f| f.material = wall_mat if f.valid? }
      end
    end
    
    model.commit_operation
    UI.messagebox("Floorplan imported!\n\nView: Camera > Standard Views > Top\nThen: Camera > Parallel Projection")
    puts "Done!"
  end
end

FloorplanImport.run
