# 2D Floorplan - Right Side Up
# Run in Ruby Console: load '/Users/vega/CascadeProjects/windsurf-project/output/floorplan_2d_fixed.rb'

module Floorplan2D
  def self.run
    model = Sketchup.active_model
    model.start_operation('2D Floorplan', true)
    
    entities = model.entities
    group = entities.add_group
    gents = group.entities
    
    def self.m(meters)
      meters * 39.3701
    end
    
    # Flip Y axis so floorplan appears right-side up
    # In SketchUp: X right, Y up (screen), Z toward viewer
    # Our data: X right, Z down (depth)
    def self.y(z_coord)
      # Flip Z so it goes up on screen instead of down
      -z_coord * 39.3701
    end
    
    mats = model.materials
    room_color = mats.add('RoomFill')
    room_color.color = Sketchup::Color.new(220, 220, 230)
    
    puts "Drawing 2D floorplan (right side up)..."
    
    # === ROOMS ===
    rooms = [
      [0, 0, 5.53, 4.0],
      [0, 4.0, 3.66, 3.06],
      [0, 7.06, 3.66, 3.3],
      [0, 10.36, 1.15, 1.5],
      [0, 11.86, 1.15, 2.22],
      [3.66, 4.0, 1.37, 5.82],
      [1.15, 10.36, 2.51, 3.72],
      [3.66, 9.82, 7.68, 5.58],
      [11.34, 9.82, 3.1, 5.58],
      [18.44, 3.12, 5.26, 3.87],
      [18.44, 6.99, 5.26, 5.29],
    ]
    
    rooms.each do |x, z, w, d|
      pts = [
        [m(x), y(z), 0],
        [m(x + w), y(z), 0],
        [m(x + w), y(z + d), 0],
        [m(x), y(z + d), 0]
      ]
      face = gents.add_face(pts)
      face.material = room_color if face
    end
    
    # === OUTER WALLS ===
    outer_walls = [
      [0, 0, 5.53, 0],
      [5.53, 0, 5.53, 4],
      [5.53, 4, 14.44, 4],
      [14.44, 4, 14.44, 15.4],
      [14.44, 15.4, 3.66, 15.4],
      [3.66, 15.4, 3.66, 14.08],
      [3.66, 14.08, 0, 14.08],
      [0, 14.08, 0, 0],
      [18.44, 3.12, 23.7, 3.12],
      [23.7, 3.12, 23.7, 12.28],
      [23.7, 12.28, 18.44, 12.28],
      [18.44, 12.28, 18.44, 3.12],
    ]
    
    outer_walls.each do |x1, z1, x2, z2|
      gents.add_line([m(x1), y(z1), 0], [m(x2), y(z2), 0])
    end
    
    # === INTERIOR WALLS ===
    interior_walls = [
      [0.8, 4, 3.66, 4],
      [3.66, 4, 3.66, 4.4],
      [3.66, 5.2, 3.66, 7.06],
      [0, 7.06, 3.66, 7.06],
      [3.66, 7.06, 3.66, 7.46],
      [3.66, 8.26, 3.66, 10.36],
      [0, 10.36, 3.66, 10.36],
      [1.15, 10.36, 1.15, 11.06],
      [1.15, 11.86, 1.15, 14.08],
      [0, 11.86, 1.15, 11.86],
      [0, 14.08, 1.15, 14.08],
      [3.66, 4, 3.66, 9.82],
      [3.66, 9.82, 5.03, 9.82],
      [3.66, 11.86, 3.66, 15.4],
      [11.34, 9.82, 11.34, 15.4],
      [18.44, 6.99, 23.7, 6.99],
    ]
    
    interior_walls.each do |x1, z1, x2, z2|
      gents.add_line([m(x1), y(z1), 0], [m(x2), y(z2), 0])
    end
    
    model.commit_operation
    
    # Set top view
    view = model.active_view
    camera = view.camera
    camera.set([m(12), y(8), m(30)], [m(12), y(8), 0], [0, -1, 0])
    camera.perspective = false
    view.zoom_extents
    view.invalidate
    
    UI.messagebox("2D Floorplan created right-side up!")
    puts "Done!"
  end
end

Floorplan2D.run
