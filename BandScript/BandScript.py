import urllib.request
import csv
import xlsxwriter

songs = ["#10 Secerisul", "#100 Pescarul", "#102 Straluceste In Departare", "#103 Golgota", "#119 Golgota", "#13 Auzi Goarna", "#16 Marsul Arad", "#2 Curatit Esti De Pacat", "#23 Privegheati", "#3 Statornic", "#55", "#56", "#66 In Noaptea Lina", "#77 Privegheati", "#78 Bucurati-va In Domnul", "#85 In Veci Cinta-Voi", "A Batut La Usa Ta", "A Lui Sa Fie Gloria", "Al Meu Pastor", "Aleluia", "Aleluia Iti Cant", "A-nviat", "As Vrea Sa Zbor", "Asculta Doamne", "Auzi Din Cer Cum Suna", "Avem Victorie", "Azi Sa Nascut", "Calea Sfanta", "Cand Eram Pandit De-o Ispita", "Cand Oceanele", "Cantati Toti De Bucurie - O Ce Veste Imbucuratoare", "Ce Minunat", "Ce Minune Va Fi", "Change My Heart", "Cinta Inima", "Cintarea Mea E Numai Despre Isus", "Cintati Toti De Bucurie - O Ce Veste Imbucuratoare", "Cu Isus In Lumea Asta", "Dealul Capatanii", "Deasupra Stelelor", "Din Cuvant Eu Am Aflat", "Din El Prin El", "Doresc Sa Fiu Aproape", "Dragoste Divina Sfanta", "Dragostea", "Dragostea Si Pacea", "Duhul Domnului", "Dumnezeu E Taria Mea", "Dupa Ploaie Si Furtuna", "E Scris Pe Calvar", "Ecoul Muntilor", "Emanuel", "Emmanuel", "Eram Jos In Pacat", "Eu Am O Calauza", "Eu Am O Carte", "Eu Sunt Alfa Si Omega", "Evanghelia (Kovacs)", "Evanghelia (Morosan)", "Evanghelia Schimba Inima", "Evanghelia Schimba Viata", "Father I Adore You", "Fii Binecuvantat Isuse", "Give Thanks", "Gloria Golgotei", "Glorie", "Glorii Din Cer", "Gratia Lui Isus", "Hai Cu Mine", "Ierusalime", "Imnul Credintei", "Imparatul Pacii", "Intoarce-te La Dragostea Dintai", "Intr-un Sat Ne-nsemnat", "Isus E Dom", "Isus Este Acelasi", "Isus Iubit (Kovacs)", "Isus Iubit (Maghiari)", "Isuse Pentru Mine", "La Calvar", "Las Sa Intre Soarele", "Leul Din Iuda", "Lumina Dulce", "Mai Sus De Toate", "Majesty", "Marea Vietii", "Maria Te-ai Gandit", "Mary Did You Know", "Noi Pentru El", "Nu Lasa Sa Treaca Vremea", "Numai Harul", "O Ce Veste Minunata", "O Holy Night", "O Isuse Domnul Meu", "O Noapte - O Ce Veste Imbucuratoare", "O Noapte Sfanta O Ce Veste Imbucuratoare", "Oda Bucuriei", "Padurea Libanului", "Paradis", "Pavel In Temnita", "Pe Bratul Domnului", "Pe Drumul Vietii Azi", "Pentru Mine Pentru Tine", "Poarta Cerurilor", "Ratacit Eram", "Regina Mea", "Rugaciunea Mamei", "Sa Fie Pace", "Sa Fiti Uniti", "Sculati Cantati", "Sfant E Numele Tau", "Silent Night", "Sonic Poems", "Sub Ochiul Sfant", "Suna Harfa", "Sunt Un Pribeag", "Sus Voi Fii", "Tara Mea", "Te-Aleg Mereu", "Te-Astept Isuse", "Traim Vremi De Har", "Trimbita Domnului", "Tu Sa Domnesti", "Umbland Isus", "Valurile Minunate", "Venirea Domnului", "Veniti Crestin", "Vesel Eram", "Vestea Buna", "Via Dolorosa", "Viata Ta", "Vin La Isus", "Vino Isus La Nunta", "Vino La Apa Vietii", "Vino Vino Tu La Isus", "Vreau Doamne", "When The Spirit Of The Lord", "Yuletide Echoes", "Zidurile Ierihonului", "Ziua E Aproape"]

dates = [None] * len(songs)

freqs = [0] * len(songs)

urllib.request.urlretrieve("https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=11Fv9xniyG2l74Bliui54Vt7WTJJ1YLMVtZaTMTU39a8&exportFormat=csv&gid=0", "temp1.csv")
print("Downloaded 1st csv file")

urllib.request.urlretrieve("https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=11Fv9xniyG2l74Bliui54Vt7WTJJ1YLMVtZaTMTU39a8&exportFormat=csv", "temp2.csv")
print("Downloaded 2nd csv file")

with open("temp1.csv") as csvf:
    reader = csv.reader(csvf, delimiter=',')

    for row in reader:

        date = row[1]

        print(date)

        for column in row:
            i = 0
            for song in songs:
                if song in column:
                    freqs[i] += 1
                    dates[i] = date
                i += 1

with open("temp2.csv") as csvf:
    reader = csv.reader(csvf, delimiter=',')

    for row in reader:

        date = row[1]

        print(date)

        for column in row:
            i = 0
            for song in songs:
                if song in column:
                    freqs[i] += 1
                    dates[i] = date
                    '''
                else:
                    songs.append(column)
                    dates.append(date)
                    '''
                i += 1


workbook = xlsxwriter.Workbook("Band Song History.xlsx")

ws = workbook.add_worksheet()

ws.write(1, 1, "Frequency")
ws.write(1, 2, "Song")
ws.write(1, 3, "Date")

row = 2

for freq, song, date in zip(freqs, songs, dates):
    print(str(date) + " " + song)
    ws.write(row, 1, freq)
    ws.write(row, 2, song)
    ws.write(row, 3, date)

    row += 1

workbook.close()