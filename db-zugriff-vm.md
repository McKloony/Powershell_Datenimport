# Datenbankzugriff aus der VMware-VM

Stand: 2026-04-27

Die VM erreicht den Windows-Host ueber die VMware-NAT-Adresse:

```text
Host-IP: 192.168.198.1
```

Nicht `localhost` verwenden. `localhost` zeigt innerhalb der VM auf die VM selbst, nicht auf den Windows-Host.

## Microsoft SQL Server

Getestet und erreichbar:

```text
Host:     192.168.198.1
Port:     1433
Server:   tcp:192.168.198.1,1433
Login:    sa
Passwort: !nic7774
Fokus-DB: Dummy
```

Getestete Server-Version:

```text
Microsoft SQL Server 2022 Express
```

Connection-String:

```text
Server=192.168.198.1,1433;Database=Dummy;User Id=sa;Password=!nic7774;Encrypt=True;TrustServerCertificate=True;
```

Beispiel fuer .NET / C#:

```csharp
var connectionString =
    "Server=192.168.198.1,1433;Database=Dummy;User Id=sa;Password=!nic7774;Encrypt=True;TrustServerCertificate=True;";
```

## MariaDB

Getestet und erreichbar:

```text
Host:     192.168.198.1
Port:     3306
Login:    root
Passwort: !nic7774
Fokus-DB: simplimed
```

Getestete Server-Version:

```text
12.3.1-MariaDB
```

Connection-String:

```text
Server=192.168.198.1;Port=3306;Database=simplimed;Uid=root;Pwd=!nic7774;
```

Beispiel fuer .NET / C#:

```csharp
var connectionString =
    "Server=192.168.198.1;Port=3306;Database=simplimed;Uid=root;Pwd=!nic7774;";
```

## Hinweise

- Die TCP-Verbindung wurde aus der VM erfolgreich getestet.
- SQL Server wurde erfolgreich mit `sa` verbunden.
- MariaDB wurde erfolgreich mit `root` verbunden.
- In der VM sind aktuell keine CLI-Clients wie `sqlcmd`, `mysql` oder `mariadb` installiert.
- Fuer produktive Nutzung besser eigene Datenbankbenutzer mit eingeschraenkten Rechten anlegen.
- Da das Passwort hier im Klartext dokumentiert ist, sollte die Datei nicht in ein Repository oder in geteilte Ordner uebernommen werden.
