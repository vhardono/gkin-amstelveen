# 2D Floorplan Only - Flat Drawing (No 3D Walls)
# Run in Ruby Console: load '/Users/vega/CascadeProjects/windsurf-project/output/floorplan_2d_only.rb'

module Floorplan2D
  def self.run
    model = Sketchup.active_model
    model.start_operation('2D Floorplan', true)
    
    entities = model.entities
    
    # Create group for all lines
    group = entities.add_group
    gents = group.entities
    
    # Convert meters to inches
    def self.m(meters)
      meters * 39.3701
    end
    
    # Colors
    mats = model.materials
    wall_color = mats.add('WallLine')
    wall_color.color = Sketchup::Color.new(100, 100, 120)
    
    room_color = mats.add('RoomFill')
    room_color.color = Sketchup::Color.new(200, 200, 210)
    
    puts "Drawing 2D floorplan..."
    
    # === ROOMS (as flat rectangles) ===
    rooms = [
      # [x, y, width, depth, name]
      [0, 0, 5.53, 4.0, "Slaapkamer 1"],
      [0, 4.0, 3.66, 3.06, "Badkamer"],
      [0, 7.06, 3.66, 3.3, "Slaapkamer 2"],
      [0, 10.36, 1.15, 1.5, "Toilet"],
      [0, 11.86, 1.15, 2.22, "CV"],
      [3.66, 4.0, 1.37, 5.82, "Foyer"],
      [1.15, 10.36, 2.51, 3.72, "Entrance"],
      [3.66, 9.82, 7.68, 5.58, "Keuken"],
      [11.34, 9.82, 3.1, 5.58, "Woonkamer"],
      [18.44, 3.12, 5.26, 3.87, "Office"],
      [18.44, 6.99, 5.26, 5.29, "Garage"],
    ]
    
    rooms.each do |x, y, w, d, name|
      # Draw rectangle outline
      pts = [
        [m(x), m(y), 0],
        [m(x + w), m(y), 0],
        [m(x + w), m(y + d), 0],
        [m(x), m(y + d), 0]
      ]
      
      # Add face for the room (filled)
      face = gents.add_face(pts)
      if face
        face.material = room_color
        face.back_material = room_color
      end
      
      # Add edges for walls (thicker lines)
      edges = []
      4.times do |i|
        edge = gents.add_line(pts[i], pts[(i + 1) % 4])
        edges << edge if edge
      end
    end
    
    # === OUTER WALL LINES (thicker) ===
    outer_walls = [
      # Main building outer shell
      [0, 0, 5.53, 0],           # Top (North) - Slaapkamer 1
      [5.53, 0, 5.53, 4],        # Right step
      [5.53, 4, 14.44, 4],       # Top (North) - Keuken area
      [14.44, 4, 14.44, 15.4],   # Right (East)
      [14.44, 15.4, 3.66, 15.4], # Bottom (South) - main
      [3.66, 15.4, 3.66, 14.08], # Step up
      [3.66, 14.08, 0, 14.08],   # Bottom (South) - left
      [0, 14.08, 0, 0],          # Left (West)
      
      # Garage/Office
      [18.44, 3.12, 23.7, 3.12], # Top
      [23.7, 3.12, 23.7, 12.28], # Right
      [23.7, 12.28, 18.44, 12.28], # Bottom
      [18.44, 12.28, 18.44, 3.12], # Left
    ]
    
    outer_walls.each do |x1, y1, x2, y2|
      gents.add_line([m(x1), m(y1), 0], [m(x2), m(y2), 0])
    end
    
    # === INTERIOR WALL LINES ===
    interior_walls = [
      # Slaapkamer 1 south
      [0.8, 4, 3.66, 4],
      # Badkamer east (two segments with door gap)
      [3.66, 4, 3.66, 4.4],
      [3.66, 5.2, 3.66, 7.06],
      # Badkamer south
      [0, 7.06, 3.66, 7.06],
      # Slaapkamer 2 east (two segments)
      [3.66, 7.06, 3.66, 7.46],
      [3.66, 8.26, 3.66, 10.36],
      # Slaapkamer 2 south
      [0, 10.36, 3.66, 10.36],
      # Toilet east (with door gap)
      [1.15, 10.36, 1.15, 11.06],
      [1.15, 11.86, 1.15, 14.08],
      # Toilet south
      [0, 11.86, 1.15, 11.86],
      # CV south
      [0, 14.08, 1.15, 14.08],
      # Foyer west
      [3.66, 4, 3.66, 9.82],
      # Foyer south
      [3.66, 9.82, 5.03, 9.82],
      # Partition wall (entrance to keuken)
      [3.66, 11.86, 3.66, 15.4],
      # Keuken/Woonkamer divider
      [11.34, 9.82, 11.34, 15.4],
      # Garage divider
      [18.44, 6.99, 23.7, 6.99],
    ]
    
    interior_walls.each do |x1, y1, x2, y2|
      gents.add_line([m(x1), m(y1), 0], [m(x2), m(y2), 0])
    end
    
    model.commit_operation
    
    # Set top view
    view = model.active_view
    camera = view.camera
    camera.set([m(12), m(8), m(30)], [m(12), m(8), 0], [0, 1, 0])
    camera.perspective = false
    view.zoom_extents
    view.invalidate
    
    UI.messagebox("2D Floorplan created!\n\nAll rooms are flat rectangles.\nUse Push/Pull to extrude walls manually.\nUse Move tool to adjust as needed.")
    puts "2D floorplan complete!"
  end
end

Floorplan2D.run
