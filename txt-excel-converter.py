import xlsxwriter

workbook = xlsxwriter.Workbook('Galaxy_007_Brigand.xlsx')
worksheet = workbook.add_worksheet()

with open('Galaxy_007_Brigand.txt', 'r') as f:
        lines = []
        for line in f:
            lines.append(line[:-1].strip())
row = 0
col = 0
curr = 0


Coordinates_x = None
Coordinates_y = None
Coordinates_z = None
Entry = None
Star_Name = None
Star_Type = None
Stellar_Mass = None
Oort_Cloud = None
Body_Name = None
Distance_from_Star = None
World_Type = None
Size = None
Diameter = None
Density = None
Gravity = None
Tilt = None
Special_Feature = None
Moonlets = None
Small_Moons = None
Medium_Moons = None
Large_Moons = None
Day_Length = None
Year_Length = None
Atmostphere_Pressure = None
Composition = None
Climate = None
Biosphere = None
Terrain_Type = None
Water = None
Humidity = None
Gemstones = None
Rare_Metals = None
Radioactives = None
Heavy_Metals = None
Industrial_Metals = None
Light_Metals = None
Organics = None

Stars=0
Planets = 0
Belts = 0

def reset():
    global Stars
    global Planets 
    global Belts 
    global Coordinates_x
    global Coordinates_y
    global Coordinates_z
    global Entry
    global Star_Name
    global Star_Type
    global Stellar_Mass
    global Oort_Cloud
    global Body_Name
    global Distance_from_Star
    global World_Type
    global Size
    global Diameter
    global Density
    global Gravity
    global Tilt
    global Special_Feature
    global Moonlets
    global Small_Moons
    global Medium_Moons
    global Large_Moons
    global Day_Length
    global Year_Length
    global Atmostphere_Pressure
    global Composition
    global Climate
    global Biosphere
    global Terrain_Type
    global Water
    global Humidity
    global Gemstones
    global Rare_Metals
    global Radioactives
    global Heavy_Metals
    global Industrial_Metals
    global Light_Metals
    global Organics

    Coordinates_x = None
    Coordinates_y = None
    Coordinates_z = None
    Entry = None
    Star_Name = None
    Star_Type = None
    Stellar_Mass = None
    Oort_Cloud = None
    Body_Name = None
    Distance_from_Star = None
    World_Type = None
    Size = None
    Diameter = None
    Density = None
    Gravity = None
    Tilt = None
    Special_Feature = None
    Moonlets = None
    Small_Moons = None
    Medium_Moons = None
    Large_Moons = None
    Day_Length = None
    Year_Length = None
    Atmostphere_Pressure = None
    Composition = None
    Climate = None
    Biosphere = None
    Terrain_Type = None
    Water = None
    Humidity = None
    Gemstones = None
    Rare_Metals = None
    Radioactives = None
    Heavy_Metals = None
    Industrial_Metals = None
    Light_Metals = None
    Organics = None
    Stars = 0
    Planets = 0
    Belts = 0

def new_body():
    global Stars
    global Planets 
    global Belts 
    global Coordinates_x
    global Coordinates_y
    global Coordinates_z
    global Entry
    global Star_Name
    global Star_Type
    global Stellar_Mass
    global Oort_Cloud
    global Body_Name
    global Distance_from_Star
    global World_Type
    global Size
    global Diameter
    global Density
    global Gravity
    global Tilt
    global Special_Feature
    global Moonlets
    global Small_Moons
    global Medium_Moons
    global Large_Moons
    global Day_Length
    global Year_Length
    global Atmostphere_Pressure
    global Composition
    global Climate
    global Biosphere
    global Terrain_Type
    global Water
    global Humidity
    global Gemstones
    global Rare_Metals
    global Radioactives
    global Heavy_Metals
    global Industrial_Metals
    global Light_Metals
    global Organics

    Entry = None
    Star_Type = None
    Stellar_Mass = None
    Oort_Cloud = None
    Body_Name = None
    Distance_from_Star = None
    World_Type = None
    Size = None
    Diameter = None
    Density = None
    Gravity = None
    Tilt = None
    Special_Feature = None
    Moonlets = None
    Small_Moons = None
    Medium_Moons = None
    Large_Moons = None
    Day_Length = None
    Year_Length = None
    Atmostphere_Pressure = None
    Composition = None
    Climate = None
    Biosphere = None
    Terrain_Type = None
    Water = None
    Humidity = None
    Gemstones = None
    Rare_Metals = None
    Radioactives = None
    Heavy_Metals = None
    Industrial_Metals = None
    Light_Metals = None
    Organics = None

def new_star():
    Star_Name = None
    Star_Type = None
    Stellar_Mass = None
    Oort_Cloud = None
    new_body()

def new_entry():
    Entry = None
    new_star()


def setup():
    row = 0
    col = 0
    content = ["X Cord", "Y Cord", "Z Cord", "Entry", "Star Name", "Star Type", "Stellar Mass", "Oort Cloud", "Body Name", "Distance from Star", "World Type", "Size", "Diameter", "Density", "Gravity", "Tilt", "Special Feature", "Moonlets", "Small Moons", "Medium Moons", "Large Moons", "Day Length", "Year Length", "Atmostphere Pressure", "Composition", "Climate", "Biosphere", "Terrain Type", "Water", "Humidity", "MR: Gemstones/ Industrial Crystals", "MR: Rare/Special Metals", "MR: Radioactives", "MR: Heavy Metals", "MR: Industrial Metals", "MR: Light Metals", "MR: Organics"]
    for item in content:
        worksheet.write(row, col, item)
        col+=1

def enter_row():
    global Coordinates_x
    global Coordinates_y
    global Coordinates_z
    global Entry
    global Star_Name
    global Star_Type
    global Stellar_Mass
    global Oort_Cloud
    global Body_Name
    global Distance_from_Star
    global World_Type
    global Size
    global Diameter
    global Density
    global Gravity
    global Tilt
    global Special_Feature
    global Moonlets
    global Small_Moons
    global Medium_Moons
    global Large_Moons
    global Day_Length
    global Year_Length
    global Atmostphere_Pressure
    global Composition
    global Climate
    global Biosphere
    global Terrain_Type
    global Water
    global Humidity
    global Gemstones
    global Rare_Metals
    global Radioactives
    global Heavy_Metals
    global Industrial_Metals
    global Light_Metals
    global Organics
    global row
    col = 0
    vars = [Coordinates_x, Coordinates_y, Coordinates_z, Entry, Star_Name, Star_Type, Stellar_Mass, Oort_Cloud, Body_Name, Distance_from_Star, World_Type, Size, Diameter, Density, Gravity, Tilt, Special_Feature, Moonlets, Small_Moons, Medium_Moons, Large_Moons, Day_Length, Year_Length, Atmostphere_Pressure, Composition, Climate, Biosphere, Terrain_Type, Water, Humidity, Gemstones, Rare_Metals, Radioactives, Heavy_Metals, Industrial_Metals, Light_Metals, Organics]
    for item in vars:
        if item!=None:
            worksheet.write(row, col, item)
        col+=1
    row +=1



def main():
    global row
    curr=0
    row = 0
    col = 0
    
    global Stars
    global Planets 
    global Belts 
    global Coordinates_x
    global Coordinates_y
    global Coordinates_z
    global Entry
    global Star_Name
    global Star_Type
    global Stellar_Mass
    global Oort_Cloud
    global Body_Name
    global Distance_from_Star
    global World_Type
    global Size
    global Diameter
    global Density
    global Gravity
    global Tilt
    global Special_Feature
    global Moonlets
    global Small_Moons
    global Medium_Moons
    global Large_Moons
    global Day_Length
    global Year_Length
    global Atmostphere_Pressure
    global Composition
    global Climate
    global Biosphere
    global Terrain_Type
    global Water
    global Humidity
    global Gemstones
    global Rare_Metals
    global Radioactives
    global Heavy_Metals
    global Industrial_Metals
    global Light_Metals
    global Organics
    percent = 0
    while curr < len(lines):
        new_percent = (round(100*(curr/len(lines))))
        if new_percent > percent:
            percent = new_percent
            print(percent,"%")
    
        line = lines[curr]
        if "Coordinates: " in line:
            enter_row()
            reset()
            coords = line.split(": ")
            coords = coords[1].split(",")
            Coordinates_x = int(coords[0])
            Coordinates_y = int(coords[1])
            Coordinates_z = int(coords[2])  
            
        if "Unusual Stellar Object:" in line and "Very" not in line:
            new_entry()
            # print(line)
            split = line.split(": ")
            Entry = split[1][:-1]
            enter_row()  

        if "Very Unusual Stellar Object:" in line:
            curr+=1
            line = lines[curr][:-1]
            Entry = line
            enter_row()   

        if "Unusual System Phenomena" in line:
            curr+=1
            line = lines[curr][:-1]
            Entry = line
            enter_row()       

        if ("Star System:" in line or "Star #" in line):
            new_entry()
            "New Entry"
            Entry = "Star System" 
            Stars+=1
            name = ""
            name += f"{Coordinates_x:0>02}{Coordinates_y:0>02}{Coordinates_z:0>02}-{chr(64+Stars)}"
            Star_Name = name

            if "Star #" not in line:
                
                curr+=1
                line = lines[curr]
                if "Star #" in line:
                    split = line.split(": ", 1)
                    Star_Type = split[1][:-1]
                else:
                    Star_Type = line[:-1]
            else:
                split = line.split(": ", 1)
                Star_Type = split[1][:-1]
            
        if "Stellar Mass:" in line:
            split = line.split(": ") 
            Stellar_Mass = float(split[1])

        if "Oort Cloud" in line:
            split = line.split(": ") 
            Oort_Cloud = split[1][:-1]

        if "Asteroid belt." in line:
            enter_row()
            new_body()
            Entry = "Asteroid Belt"
            Belts+=1
            Body_Name = f"{Star_Name}-AB{Belts:02d}"

        if "Planet #" in line:
            enter_row()
            new_body()
            Entry = "Planet"
            Planets+=1
            Body_Name = f"{Star_Name}-P{Planets:02d}"
            
        if "Distance from star:" in line:
            split = line.split(": ") 
            Distance_from_Star = split[1][:-1]
            
        if "Gas giant" in line:
            World_Type = "Gas Giant"
            if "Size: " in line:
                split = line.split(": ")
                Size = split[1][:-1]
            
        if "Terrestrial-type" in line:
            World_Type = "Terrestrial"

        if "Diameter:" in line:
            split = line.split(": ") 
            Diameter = split[1][:-1]
                  
        if "Density:" in line:
            split = line.split(": ") 
            Density = float(split[1][:-1])
            
        if "Gravity:" in line:
            split = line.split(": ") 
            Gravity = split[1]
            
        if "Axial Tilt:" in line:
            split = line.split(": ") 
            Tilt = split[1][:-1]
            
        if "Special Feature" in line:
            split = line.split(": ") 
            Special_Feature = split[1][:-1]
            
        if "Moonlets" in line:
            split = line.split(" ") 
            Moonlets = int(split[0])
            
        if "Small Moons" in line:
            split = line.split(" ") 
            Small_Moons = int(split[0])
            
        if "Medium Moons" in line:
            split = line.split(" ") 
            Medium_Moons = int(split[0])
            
        if "Large Moons" in line:
            split = line.split(" ") 
            Large_Moons = int(split[0])
                 
        if "Length of Day:" in line:
            split = line.split(": ") 
            Day_Length = split[1][:-1]
            
        if "Length of Year:" in line:
            split = line.split(": ") 
            Year_Length = split[1][:-1]
            
        if "Atmosphere Pressure:" in line:
            split = line.split(": ") 
            Atmostphere_Pressure = split[1][:-1]
            
        if "Composition:" in line:
            split = line.split(": ") 
            Composition = split[1][:-1]
            
        if "Biosphere:" in line:
            split = line.split(": ") 
            Biosphere = split[1][:-1]
            
        if "Terrain Type:" in line:
            split = line.split(": ") 
            Terrain_Type = split[1][:-1]
                  
        if "Surface Water:" in line:
            split = line.split(": ") 
            Water = split[1][:-1]
        
        if "Humidity:" in line:
            split = line.split(": ") 
            if split[1].endswith("."):
                Humidity = split[1][:-1]
            else:
                Humidity = split[1]
            
        if "Gemstones/Industrial Crystals:" in line:
            split = line.split(": ") 
            Gemstones = split[1][:-1]

        if "Rare/Special Minerals:" in line:
            split = line.split(": ") 
            Rare_Metals = split[1][:-1]

        if "Radioactives:" in line:
            split = line.split(": ") 
            Radioactives = split[1][:-1]

        if "Heavy Metals:" in line:
            split = line.split(": ") 
            Heavy_Metals = split[1][:-1]

        if "Industrial Metals:" in line:
            split = line.split(": ") 
            Industrial_Metals = split[1][:-1]

        if "Light Metals:" in line:
            split = line.split(": ") 
            Light_Metals = split[1][:-1]

        if "Organics:" in line:
            split = line.split(": ") 
            Organics = split[1][:-1]

        curr+=1
    enter_row()


setup()
main()
print("done")
f.close()
workbook.close()
