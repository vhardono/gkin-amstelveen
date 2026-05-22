# SketchUp Ruby Script - Import Floorplan from JSON
# Run this in SketchUp: Window > Ruby Console, then paste and run

require 'json'

# Path to your JSON file - update this if needed
json_path = '/Users/vega/CascadeProjects/windsurf-project/output/floorplan-data.json'

# Read and parse JSON
json_data = File.read(json_path)
data = JSON.parse(json_data)

# Create a new group for the floorplan
model = Sketchup.active_model
entities = model.entities
group = entities.add_group
grp_entities = group.entities

# Materials
wall_material = grp_entities.parent.materials.add('Wall')
wall_material.color = Sketchup::Color.new(240, 235, 228)  # #f0ebe4

glass_material = grp_entities.parent.materials.add('Glass')
glass_material.color = Sketchup::Color.new(136, 204, 238)  # #88ccee
glass_material.alpha = 0.35

floor_material = grp_entities.parent.materials.add('Floor')
floor_material.color = Sketchup::Color.new(74, 74, 74)  # #4a4a4a

# Helper to convert meters to inches (SketchUp internal units)
def m_to_in(meters)
  meters * 39.3701
end

# Wall thickness and height
wall_thickness = m_to_in(0.2)
wall_height = m_to_in(2.8)
garage_height = m_to_in(2.46)

puts "Creating floorplan..."

# Create room floors from main building
data['main_building']['rooms'].each do |room|
  x1 = m_to_in(room['coordinates']['x1'])
  x2 = m_to_in(room['coordinates']['x2'])
  z1 = m_to_in(room['coordinates']['z1'])
  z2 = m_to_in(room['coordinates']['z2'])
  
  width = x2 - x1
  depth = z2 - z1
  
  # Create floor face
  pt1 = [x1, 0, z1]
  pt2 = [x2, 0, z1]
  pt3 = [x2, 0, z2]
  pt4 = [x1, 0, z2]
  
  face = grp_entities.add_face(pt1, pt2, pt3, pt4)
  face.material = floor_material if face
  
  puts "Created floor: #{room['name']}"
end

# Create garage/office floors
data['garage_office']['rooms'].each do |room|
  x1 = m_to_in(room['coordinates']['x1'])
  x2 = m_to_in(room['coordinates']['x2'])
  z1 = m_to_in(room['coordinates']['z1'])
  z2 = m_to_in(room['coordinates']['z2'])
  
  pt1 = [x1, 0, z1]
  pt2 = [x2, 0, z1]
  pt3 = [x2, 0, z2]
  pt4 = [x1, 0, z2]
  
  face = grp_entities.add_face(pt1, pt2, pt3, pt4)
  face.material = floor_material if face
  
  puts "Created floor: #{room['name']}"
end

# Create walls from outer walls
data['walls']['outer'].each do |wall|
  h = wall['height'] ? m_to_in(wall['height']) : wall_height
  
  if wall['x'] # Vertical wall (x constant)
    x = m_to_in(wall['x'])
    z1 = m_to_in(wall['z1'])
    z2 = m_to_in(wall['z2'])
    
    pt1 = [x - wall_thickness/2, 0, z1]
    pt2 = [x + wall_thickness/2, 0, z1]
    pt3 = [x + wall_thickness/2, 0, z2]
    pt4 = [x - wall_thickness/2, 0, z2]
    
    face = grp_entities.add_face(pt1, pt2, pt3, pt4)
    if face
      face.pushpull(-h)
      face.material = wall_material
    end
  else # Horizontal wall (z constant)
    x1 = m_to_in(wall['x1'])
    x2 = m_to_in(wall['x2'])
    z = m_to_in(wall['z'])
    
    pt1 = [x1, 0, z - wall_thickness/2]
    pt2 = [x2, 0, z - wall_thickness/2]
    pt3 = [x2, 0, z + wall_thickness/2]
    pt4 = [x1, 0, z + wall_thickness/2]
    
    face = grp_entities.add_face(pt1, pt2, pt3, pt4)
    if face
      face.pushpull(-h)
      face.material = wall_material
    end
  end
end

# Create interior walls
data['walls']['interior'].each do |wall|
  h = wall['height'] ? m_to_in(wall['height']) : wall_height
  
  if wall['x'] # Vertical wall
    x = m_to_in(wall['x'])
    z1 = m_to_in(wall['z1'])
    z2 = m_to_in(wall['z2'])
    
    pt1 = [x - wall_thickness/2, 0, z1]
    pt2 = [x + wall_thickness/2, 0, z1]
    pt3 = [x + wall_thickness/2, 0, z2]
    pt4 = [x - wall_thickness/2, 0, z2]
    
    face = grp_entities.add_face(pt1, pt2, pt3, pt4)
    face.pushpull(-h) if face
  else # Horizontal wall
    x1 = m_to_in(wall['x1'])
    x2 = m_to_in(wall['x2'])
    z = m_to_in(wall['z'])
    
    pt1 = [x1, 0, z - wall_thickness/2]
    pt2 = [x2, 0, z - wall_thickness/2]
    pt3 = [x2, 0, z + wall_thickness/2]
    pt4 = [x1, 0, z + wall_thickness/2]
    
    face = grp_entities.add_face(pt1, pt2, pt3, pt4)
    face.pushpull(-h) if face
  end
end

puts "Floorplan created successfully!"
puts "Use Camera > Standard Views > Top to see the plan"
puts "Use Camera > Parallel Projection for 2D view"
