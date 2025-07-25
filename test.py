import win32com.client




xlapp = win32com.client.Dispatch("Excel.Application")
wb = xlapp.Workbooks.Open('https://pucpredu.sharepoint.com/:x:/r/teams/ProjetodeEngenharia401/Shared%20Documents/General/Entregas%20de%20grupo/4%20-%20Projeto%20Conceitual/An%C3%A1lise%20de%20fun%C3%A7%C3%B5es%20da%20Bota%20Pneum%C3%A1tica.xlsx?d=w86f6be0a5b2e4fe9af67fc2e4e9eafb5&csf=1&web=1&e=IvnU9d')