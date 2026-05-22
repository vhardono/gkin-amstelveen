# Floorplan Import Script - Horizontal Layout (Top-Down View)
# Run in Ruby Console: load '/Users/vega/CascadeProjects/windsurf-project/output/floorplan_horizontal.rb'

module FloorplanImport
  def self.run
    model = Sketchup.active_model
    model.start_operation('Import Floorplan', true)
    
    entities = model.entities
    
    # Create main group
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
    wall_h = m(2.8)
    
    puts "Creating floorplan on XY plane (horizontal)..."
    
    # Create floors - using X and Y (Z will be 0, floor is flat on ground)
    # In SketchUp: X = width, Y = depth, Z = height
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
    
    floors.each do |x1, y1, w, d|
      pts = [
        [m(x1), m(y1), 0],
        [m(x1 + w), m(y1), 0],
        [m(x1 + w), m(y1 + d), 0],
        [m(x1), m(y1 + d), 0]
      ]
      face = gents.add_face(pts)
      face.material = floor_mat if face && face.valid?
    end
    
    puts "Creating walls (pulling up in Z)..."
    
    # Create walls as lines that get pulled up (extruded in Z)
    walls_group = gents.add_group
    w_ents = walls_group.entities
    
    # Horizontal walls (constant Y, varying X) - extrude in Z
    h_walls = [
      [0, 5.53, 0, 2.8],
      [5.03, 14.44, 4.0, 2.8],
      [3.66, 14.44, 15.4, 2.8],
      [0, 3.66, 14.08, 2.8],
      [18.44, 23.70, 3.12, 2.8],
      [18.44, 23.70, 12.28, 2.46],
    ]
    
    h_walls.each do |x1, x2, y, h|
      # Create wall as a face at the position, then push up in Z
      wg = w_ents.add_group
      we = wg.entities
      
      pts = [
        [m(x1), m(y) - thick/2, 0],
        [m(x2), m(y) - thick/2, 0],
        [m(x2), m(y) + thick/2, 0],
        [m(x1), m(y) + thick/2, 0]
      ]
      face = we.add_face(pts)
      if face && face.valid?
        face.pushpull(m(h))
        we.grep(Sketchup::Face).each { |f| f.material = wall_mat if f.valid? }
      end
    end
    
    # Vertical walls (constant X, varying Y) - extrude in Z
    v_walls = [
      [0, 0, 14.08, 2.8],
      [14.44, 4.0, 15.4, 2.8],
      [18.44, 3.12, 12.28, 2.46],
      [23.70, 3.12, 12.28, 2.46],
    ]
    
    v_walls.each do |x, y1, y2, h|
      wg = w_ents.add_group
      we = wg.entities
      
      pts = [
        [m(x) - thick/2, m(y1), 0],
        [m(x) + thick/2, m(y1), 0],
        [m(x) + thick/2, m(y2), 0],
        [m(x) - thick/2, m(y2), 0]
      ]
      face = we.add_face(pts)
      if face && face.valid?
        face.pushpull(m(h))
        we.grep(Sketchup::Face).each { |f| f.material = wall_mat if f.valid? }
      end
    end
    
    model.commit_operation
    
    # Auto-switch to top view
    view = model.active_view
    camera = view.camera
    camera.set([m(12), m(8), m(40)], [m(12), m(8), 0], [0, 1, 0])
    camera.perspective = false
    view.invalidate
    
    UI.messagebox("Floorplan imported horizontally!\n\nYou should see a top-down view now.\nUse mouse wheel to zoom, Shift+drag to pan.")
    puts "Done! Floor is on XY plane, walls go up in Z."
  end
end

FloorplanImport.run
