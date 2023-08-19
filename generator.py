import random
import openpyxl

workbook = openpyxl.Workbook()
worksheet = workbook.active
num_rows = 1000 # total number of rows
header_row = ['Name', 'Age', 'Gender', 'Province', 'Municipal', 'Cause of death', 'Date of Death', 'Funeral Service']

worksheet.append(header_row)

listOfFirstName = [
    ["James", "Daniel", "David", "Thomas", "Henry", "Oliver", "Benjamin", "Jack", "Noah", "John Michael", "John Gabriel", "Mathew Lucas", "Luke", "Anthony", "Robert", "Andrew", "Juan", "Anthony", "Josh", "Martin", "Nate", "Christian"],
    ["Emma", "Anna", "Alice", "Sarah", "Emily", "Ella",  "Maria", "Olivia", "Evelyn", "Grace", "Marie Jane", "Jennifer", "Mae Ann", "Isabella", "Marry Joy", "Ella Mae", "Patricia", "Susane", "Mia", "Christine", "Amelia"],
    ["Cruz",  "Reyes", "Santos", "Garcia", "Torres", "Del Rosario", "Perez", "Fernandez", "Ramos", "Aquino", "Castro", "Martinez", "Navarro", "Sanchez", "Dela Cruz", "Salazar", "Villanueva", "Ramirez", "Castillo", "Flores", "Rivera", "Manalo", "Dela Cruz", "Mendoza", "Gonzales", "de Leon", "Martinez", "Diaz", "Hernandez",  "Vargas", "Abad", "Robles", "Conception",  "Galang", "Villa"]

]

lname = ["Cruz",  "Reyes", "Santos", "Garcia", "Torres", "Del Rosario", "Perez", "Fernandez", "Ramos", "Aquino", "Castro", "Martinez", "Navarro", "Sanchez", "Dela Cruz", "Salazar", "Villanueva", "Ramirez", "Castillo", "Flores", "Rivera", "Manalo", "Dela Cruz", "Mendoza", "Gonzales", "de Leon", "Martinez", "Diaz", "Hernandez",  "Vargas", "Abad", "Robles", "Conception",  "Galang", "Villa"]

def name_gender_Generator ():
    
    numMorF = random.randint(0, 1)
    f_name = random.choice(listOfFirstName[numMorF])
    l_name = random.choice(lname)
    name = f_name + " " + l_name
    #age = random.randint(5, 95)
    
    if numMorF == 0:
        gender = 'Male'
    else:
        gender = 'Female'

    return name, gender

def age_cause_generator(cs, limit, count_limit):
    flag = True
    cause_of_death = 'asd'
    while flag:
        flag1 = False
        flag2 = False
        start_random = 5
        if (flag1 and flag2):
            start_random = 88
        
        end_random = 95
        age = random.randint(start_random, end_random)
        if age > 88:
            cause = 'Natural Death'
            if count_limit[2] == limit[2]:
                end_random = 88
            else:
                count_limit[2] += 1
                cause_of_death = cause
                flag = False
        else:
            cause = random.choice(cs)
            if cause == cs[0]:
                c_index = cs.index(cause)
                if count_limit[c_index] == limit[c_index]:
                    flag1 = True
                    flag2 = True
                    continue
                else:
                    count_limit[0] += 1
                    cause_of_death = cause
                    flag = False
                
            if cause == cs[1]:
                c_index1 = cs.index(cause)
                if count_limit[c_index1] == limit[c_index1]:
                    flag2 = True
                    flag1 = True
                    continue
                else:
                    count_limit[1] += 1
                    cause_of_death = cause
                    flag = False

    return age, cause_of_death 
                
causes = ['Health Condition','Accident'] 
limit_cause = [500, 200, 300] #total how many dies from health condition, accident, old age
count_limit_cause = [0,0,0]

for i in range(num_rows):

    name, gender = name_gender_Generator()
    
    age, cause = age_cause_generator(causes, limit_cause, count_limit_cause)

    provinces = 'Laguna'

    municipals = ["Alaminos", "Bay", "Biñan", "Cabuyao", "Calamba", "Calauan", "Cavinti", "Famy", "Kalayaan", "Liliw", "Los Baños", "Luisana", "Lumban", "Mabitac", "Magdalena", "Majayjay", "Nagcarlan", "Nagcarlan", "Paete", "Pagsanjan", "Pakil", "Pangil", "Pila", "Rizal", "San Pablo", "San Pedro", "Santa Cruz", "Santa Maria", "Santa Rosa", "Siniloan", "Victoria"]
    municipal = random.choice(municipals)

    # date starts at January 2023
    day = random.randint(1, 31)
    date_un = '2023-01-'
    if day < 10:
        string_day = str(day)
        date = date_un + '0' + string_day
    else:
        string_day = str(day)
        date = date_un + string_day
    
    services = ["Balubayan", "Camargo", "St. Peter", "Susan Vasquez Funeral", "Furinaria Regina", "Mejai Funeral Homes", "Lescano Funeral Homes", "La Forinaria Allen Surbano", "Kapunitan Funeral Homes", "King of Heavens Funeral Service", "City Chapel", "De Mesa Funeral Homes", "Blessed Mary Funeral Homes", "Nila Belen Funeral Homes", "Rodel Señerez Funeral Sevices", "Zuasola Funeral Homes", " St. Angela Funeral Homes", "Nazareno Memorial Services", "Ace Funeral Homes", "John Paul Funeral Homes"]
    service = random.choice(services)

    row = [name, age, gender, provinces, municipal, cause, date, service]
    worksheet.append(row)

workbook.save('data.xlsx') # path and file name
print("Done")
