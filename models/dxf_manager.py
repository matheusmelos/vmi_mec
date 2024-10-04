import ezdxf
import math

class DXF:
    def __init__(self, dxf_file, dxf_name):
        
        self.dxf_file = dxf_file
        self.dxf_name = dxf_name
        self.material = ' '
        self.cut_speed = 167
        self.perimeter = DXF.calculate_perimeter_total(self)
        self.cut_time  = DXF.calculate_cut_time(self)
        
    def calculate_perimeter_total(self):
                
        doc = ezdxf.readfile(self.dxf_file)
        modelspace = doc.modelspace()
        
        lines = modelspace.query('LINE')
        arcs = modelspace.query('ARC')
        circles = modelspace.query('CIRCLE')
        lwpolylines = modelspace.query('LWPOLYLINE')
        ellipses = modelspace.query('ELLIPSE')
        polylines = modelspace.query('POLYLINE')
        
        perimeter_total = (
            DXF.calculate_perimeter_of_lines(lines) +
            DXF.calculate_perimeter_of_arcs(arcs) +
            DXF.calculate_perimeter_of_circles(circles) +
            DXF.calculate_perimeter_of_lwpolylines(lwpolylines) +
            DXF.calculate_perimeter_of_ellipses(ellipses) +
            DXF.calculate_perimeter_of_lwpolylines(polylines)
        )

        return perimeter_total
    
    def calculate_cut_time(self):
        cut_time = (self.perimeter)/(self.cut_speed)
        return f"{cut_time:.3f}"
        
    def calculate_distance(point1, point2):
        return math.sqrt((point2[0] - point1[0]) ** 2 + (point2[1] - point1[1]) ** 2)

    def calculate_perimeter_of_lines(lines):
        perimeter = 0.0
        for line in lines:
            start_point = (line.dxf.start.x, line.dxf.start.y)
            end_point = (line.dxf.end.x, line.dxf.end.y)
            perimeter += DXF.calculate_distance(start_point, end_point)

        return perimeter

    def calculate_perimeter_of_arcs(arcs):
        perimeter = 0.0
        for arc in arcs:
            start_angle = arc.dxf.start_angle
            end_angle = arc.dxf.end_angle
            radius = arc.dxf.radius
            angle_diff = abs(end_angle - start_angle)
            angle_diff = math.radians(angle_diff)
            perimeter += radius * angle_diff
        
        return perimeter

    def calculate_perimeter_of_circles(circles):
        perimeter = 0.0
        for circle in circles:
            radius = circle.dxf.radius
            perimeter += 2 * math.pi * radius

        return perimeter
 
    def calculate_perimeter_of_lwpolylines(lwpolylines):
        perimeter = 0.0
        for polyline in lwpolylines:
            vertices = list(polyline.vertices())
            
            for i in range(len(vertices) - 1):
                perimeter += DXF.calculate_distance(vertices[i], vertices[i + 1])
            
            if polyline.close: 
                perimeter += DXF.calculate_distance(vertices[-1], vertices[0])
        
        
        return perimeter

    def calculate_perimeter_of_ellipses(ellipses):
        perimeter =  0.0  
        for ellipse in ellipses: 
            a = ellipse.dxf.major_axis.magnitude / 2
            b = ellipse.dxf.minor_axis.magnitude / 2
            perimeter += math.pi * (3 * (a + b) - math.sqrt((3 * a + b) * (a + 3 * b)))
        return perimeter

