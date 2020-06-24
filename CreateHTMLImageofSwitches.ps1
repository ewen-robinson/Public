#Script
# 11/05/2020

#variables

$source = "C:\Users\xxxxxxx\.networkassistant\backups\"

$target = "C:\TempPath\"
$pattern = "startup-config"
$record = $false
$global:array = @()
$global:interface = $null

#functions
function Add-ToNewObject {
    param ($type)
    #write-host $content[$i] 
    $global:interface = New-Object -TypeName psobject
    $interface | Add-Member -MemberType NoteProperty -Name Type -Value $type
    $interface | Add-Member -MemberType NoteProperty -Name Number -Value ($content[$i] -split '/')[-1]   
}

function Add-ToExistingObject {
    #switchport access
    Switch -Wildcard ($content[$i]) {
        "*switchport access*" { $interface | Add-Member -MemberType NoteProperty -Name VLAN -Value ($content[$i] -split ' ')[-1] }
        "*switchport mode trunk*" { $interface | Add-Member -MemberType NoteProperty -Name VLAN -Value "Trunk" }
        "*switchport voice*" {$interface | Add-Member -MemberType NoteProperty -Name Voice -Value $true}
    }
}

function Close-Object {
    $global:array += $interface
}


function 24Port-HTML {
    param ($swName)
    
        #debug code
       # write-host $swName

    if (@($array | Where-Object {$_.Type -eq "g"}).Count -lt 8) {
      foreach ($result in $array) {
        switch ($result.VLAN) {
          "Trunk" {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#FF00FF"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#FF00FF"}; Break}
          10 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#7FFF00"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#7FFF00"}; if ($result.Voice) {New-Variable -Name "splitport$($result.Number)" -Value "background-image: -webkit-linear-gradient(135deg, #7FFF00 50%, #0000FF 50%);"}}
          30 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#DC143C"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#DC143C"}}
          31 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#9932CC"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#9932CC"}}
          35 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#E67E22"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#E67E22"}}
          123 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#0000FF"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#0000FF"}}
          Default {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#616A6B"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#616A6B"}}
        }
      }
    } else {
      foreach ($result in $array) {
        switch ($result.VLAN) {
          "Trunk" {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#FF00FF"}; Break}
          10 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#7FFF00"}; if ($result.Voice) {New-Variable -Name "splitport$($result.Number)" -Value "background-image: -webkit-linear-gradient(135deg, #7FFF00 50%, #0000FF 50%);"}}
          30 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#DC143C"}}
          31 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#9932CC"}}
          35 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#E67E22"}}
          123 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#0000FF"}}
          Default {if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#616A6B"}}
        }
      }
    }

    $fileName = $swName.split('_')[0]
    
    $html = @"
  <!DOCTYPE html>
  <head>
  <style>
  .grid-container {
    background-color:#eee;
    width: 710px;
    margin: 20px auto;
    
    display: grid;
    grid-template-rows: 100px 50px 50px 50px 50px 50px 50px 50px 25px;
    grid-template-columns: repeat(13, 50px) ;
  
    grid-gap: 5px; 
    
    color: #000;
    font-family: sans-serif;
  }
  
  header {
    background-color: #C2CA2F;
    grid-row: 1 / 1;
    grid-column: 1 / -1;
    color: #656365;
  }
  
  .key-box {
   color: #000;
  }
  
  .key-box--1 {
        grid-row: 5 / span 1;
        grid-column: 1 / span 2;
    }
  .key-box--2 {
        grid-row: 6 / span 1;
        grid-column: 1 / span 2;
    }
  .key-box--3 {
        grid-row: 7 / span 1;
        grid-column: 1 / span 2;
    }
  .key-box--4 {
        grid-row: 8 / span 1;
        grid-column: 1 / span 2;
    }
  .key-box--5 {
      grid-row: 5 / span 1;
      grid-column: 4 / span 2;
  }
  .key-box--6 {
    grid-row: 6 / span 1;
    grid-column: 4 / span 2;
  }
  .key-box--7 {
    grid-row: 7 / span 1;
    grid-column: 4 / span 2;
  }
  
  .g-box {
      border: 1px solid black;
  }
  .g-box--1 {
      grid-row: 2 / span 1;
      grid-column: 13 / span 1;
  }
  .g-box--2 {
      grid-row: 3 / span 1;
      grid-column: 13 / span 1;
  } 
  
  .sm-box {
    background-color: #C2CA2F;
    font-size: 15px;
   } 
  .sm-box--1 {
      grid-row: 2 / span 1;
      grid-column: 1 / span 1;
    }
  .sm-box--2 {
      grid-row: 2 / span 1;
      grid-column: 2 / span 1;
    }
  .sm-box--3 {
      grid-row: 2 / span 1;
      grid-column: 3 / span 1;
    }
  .sm-box--4 {
      grid-row: 2 / span 1;
      grid-column: 4 / span 1;
    }
  .sm-box--5 {
      grid-row: 2 / span 1;
      grid-column: 5 / span 1;
    }
  .sm-box--6 {
      grid-row: 2 / span 1;
      grid-column: 6 / span 1;
    }
  .sm-box--7 {
      grid-row: 2 / span 1;
      grid-column: 7 / span 1;
    }
  .sm-box--8 {
      grid-row: 2 / span 1;
      grid-column: 8 / span 1;
    }
  .sm-box--9 {
      grid-row: 2 / span 1;
      grid-column: 9 / span 1;
    }
  .sm-box--10 {
      grid-row: 2 / span 1;
      grid-column: 10 / span 1;
    }
  .sm-box--11 {
      grid-row: 2 / span 1;
      grid-column: 11 / span 1;
    }
  .sm-box--12 {
      grid-row: 2 / span 1;
      grid-column: 12 / span 1;
    }
  .sm-box--13 {
      grid-row: 3 / span 1;
      grid-column: 1 / span 1;
    }
  .sm-box--14 {
      grid-row: 3 / span 1;
      grid-column: 2 / span 1;
    }
  .sm-box--15 {
      grid-row: 3 / span 1;
      grid-column: 3 / span 1;
    }
   .sm-box--16 {
      grid-row: 3 / span 1;
      grid-column: 4 / span 1;
    }
  .sm-box--17 {
      grid-row: 3 / span 1;
      grid-column: 5 / span 1;
    }
  .sm-box--18 {
      grid-row: 3 / span 1;
      grid-column: 6 / span 1;
    }
  .sm-box--19 {
      grid-row: 3 / span 1;
      grid-column: 7 / span 1;
    }
  .sm-box--20 {
      grid-row: 3 / span 1;
      grid-column: 8 / span 1;
    }
   .sm-box--21 {
      grid-row: 3 / span 1;
      grid-column: 9 / span 1;
    }
  .sm-box--22 {
      grid-row: 3 / span 1;
      grid-column: 10 / span 1;
    }
  .sm-box--23 {
      grid-row: 3 / span 1;
      grid-column: 11 / span 1;
    }
  .sm-box--24 {
      grid-row: 3 / span 1;
      grid-column: 12 / span 1;
    }
  </style>
  
  <div class="grid-container">
    <header>$fileName</header>
    <div class="sm-box sm-box--1" style="background-color:$port1;$splitport1">1</div>
    <div class="sm-box sm-box--2" style="background-color:$port2;$splitport2">2</div>
    <div class="sm-box sm-box--3" style="background-color:$port3;$splitport3">3</div>
    <div class="sm-box sm-box--4" style="background-color:$port4;$splitport4">4</div>
    <div class="sm-box sm-box--5" style="background-color:$port5;$splitport5">5</div>
    <div class="sm-box sm-box--6" style="background-color:$port6;$splitport6">6</div>
    <div class="sm-box sm-box--7" style="background-color:$port7;$splitport7">7</div>
    <div class="sm-box sm-box--8" style="background-color:$port8;$splitport8">8</div>
    <div class="sm-box sm-box--9" style="background-color:$port9;$splitport9">9</div>
    <div class="sm-box sm-box--10" style="background-color:$port10;$splitport10">10</div>
    <div class="sm-box sm-box--11" style="background-color:$port11;$splitport11">11</div>
    <div class="sm-box sm-box--12" style="background-color:$port12;$splitport12">12</div>
    <div class="sm-box sm-box--13" style="background-color:$port13;$splitport13">13</div>
    <div class="sm-box sm-box--14" style="background-color:$port14;$splitport14">14</div>
    <div class="sm-box sm-box--15" style="background-color:$port15;$splitport15">15</div>
    <div class="sm-box sm-box--16" style="background-color:$port16;$splitport16">16</div>
    <div class="sm-box sm-box--17" style="background-color:$port17;$splitport17">17</div>
    <div class="sm-box sm-box--18" style="background-color:$port18;$splitport18">18</div>
    <div class="sm-box sm-box--19" style="background-color:$port19;$splitport19">19</div>
    <div class="sm-box sm-box--20" style="background-color:$port20;$splitport20">20</div>
    <div class="sm-box sm-box--21" style="background-color:$port21;$splitport21">21</div>
    <div class="sm-box sm-box--22" style="background-color:$port22;$splitport22">22</div>
    <div class="sm-box sm-box--23" style="background-color:$port23;$splitport23">23</div>
    <div class="sm-box sm-box--24" style="background-color:$port24;$splitport24">24</div>
    <div class="g-box g-box--1" style="background-color:$gport1;">G1</div>
    <div class="g-box g-box--2" style="background-color:$gport2;">G2</div>
    <div class="key-box key-box--1" style="background-color:#7FFF00;">Staff</div>
    <div class="key-box key-box--2" style="background-color:#DC143C;">Public</div>
    <div class="key-box key-box--3" style="background-color:#0000FF;">Voice</div>
    <div class="key-box key-box--4" style="background-color:#9932CC;">WIFI</div>
    <div class="key-box key-box--5" style="background-color:#E67E22;">RFID</div>
    <div class="key-box key-box--6" style="background-color:#FF00FF;">Trunk</div>
    <div class="key-box key-box--7" style="background-color:#616A6B;">Other</div>
  </div>
"@
  
  
  $html | set-content -path "c:\scripts\switch\$fileName.html"
}
  
function 48Port-HTML {
    param ($swName)
    #debug code
    #write-host $swName

    if (@($array | Where-Object {$_.Type -eq "g"}).Count -lt 8) {
      foreach ($result in $array) {
        switch ($result.VLAN) {
          "Trunk" {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#FF00FF"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#FF00FF"}; Break}
          10 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#7FFF00"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#7FFF00"}; if ($result.Voice) {New-Variable -Name "splitport$($result.Number)" -Value "background-image: -webkit-linear-gradient(135deg, #0000FF 50%, #7FFF00 51%);"}}
          30 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#DC143C"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#DC143C"}}
          31 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#9932CC"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#9932CC"}}
          35 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#E67E22"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#E67E22"}}
          123 {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#0000FF"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#0000FF"}}
          Default {if ($result.type -eq "f") {New-Variable -Name "port$($result.Number)" -Value "#616A6B"}; if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#616A6B"}}
        }
      }
    } else {
      foreach ($result in $array) {
        switch ($result.VLAN) {
          "Trunk" {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#FF00FF"}; Break}
          10 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#7FFF00"}; if ($result.Voice) {New-Variable -Name "splitport$($result.Number)" -Value "background-image: -webkit-linear-gradient(135deg, #0000FF 50%, #7FFF00 51%);"}}
          30 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#DC143C"}}
          31 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#9932CC"}}
          35 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#E67E22"}}
          123 {if ($result.type -eq "g") {New-Variable -Name "port$($result.Number)" -Value "#0000FF"}}
          Default {if ($result.type -eq "g") {New-Variable -Name "gport$($result.Number)" -Value "#616A6B"}}
        }
      }
    }

    $fileName = $swName.split('_')[0]

    $html = @"
<!DOCTYPE html>
<head>
<style>
.grid-container {
  background-color:#eee;
  width: 775px;
  margin: 20px auto;
  
  display: grid;
  grid-template-rows: 100px 25px 25px 25px 25px 25px 25px 25px 13px;
  grid-template-columns: repeat(26, 25px) ;

  grid-gap: 5px; 
  
  color: #000;
  font-family: sans-serif;
}

header {
  background-color: #C2CA2F;
  grid-row: 1 / 1;
  grid-column: 1 / -1;
  color: #656365;
}

.key-box {
 color: #000;
}

.key-box--1 {
      grid-row: 5 / span 1;
      grid-column: 1 / span 2;
  }
.key-box--2 {
      grid-row: 6 / span 1;
      grid-column: 1 / span 2;
  }
.key-box--3 {
      grid-row: 7 / span 1;
      grid-column: 1 / span 2;
  }
.key-box--4 {
      grid-row: 8 / span 1;
      grid-column: 1 / span 2;
  }
.key-box--5 {
    grid-row: 5 / span 1;
    grid-column: 4 / span 2;
}
.key-box--6 {
  grid-row: 6 / span 1;
  grid-column: 4 / span 2;
}
.key-box--7 {
  grid-row: 7 / span 1;
  grid-column: 4 / span 2;
}

.g-box {
	border: 1px solid black;
}
.g-box--1 {
    grid-row: 2 / span 1;
    grid-column: 25 / span 1;
}
.g-box--2 {
    grid-row: 3 / span 1;
    grid-column: 25 / span 1;
} 
.g-box--3 {
    grid-row: 2 / span 1;
    grid-column: 26 / span 1;
}
.g-box--4 {
    grid-row: 3 / span 1;
    grid-column: 26 / span 1;
} 

.sm-box {
  background-color: #C2CA2F;
  font-size: 15px;
 } 
.sm-box--1 {
    grid-row: 2 / span 1;
    grid-column: 1 / span 1;
  }
.sm-box--2 {
    grid-row: 2 / span 1;
    grid-column: 2 / span 1;
  }
.sm-box--3 {
    grid-row: 2 / span 1;
    grid-column: 3 / span 1;
  }
.sm-box--4 {
    grid-row: 2 / span 1;
    grid-column: 4 / span 1;
  }
.sm-box--5 {
    grid-row: 2 / span 1;
    grid-column: 5 / span 1;
  }
.sm-box--6 {
    grid-row: 2 / span 1;
    grid-column: 6 / span 1;
  }
.sm-box--7 {
    grid-row: 2 / span 1;
    grid-column: 7 / span 1;
  }
.sm-box--8 {
    grid-row: 2 / span 1;
    grid-column: 8 / span 1;
  }
.sm-box--9 {
    grid-row: 2 / span 1;
    grid-column: 9 / span 1;
  }
.sm-box--10 {
    grid-row: 2 / span 1;
    grid-column: 10 / span 1;
  }
.sm-box--11 {
    grid-row: 2 / span 1;
    grid-column: 11 / span 1;
  }
.sm-box--12 {
    grid-row: 2 / span 1;
    grid-column: 12 / span 1;
  }
.sm-box--13 {
    grid-row: 2 / span 1;
    grid-column: 13 / span 1;
  }
.sm-box--14 {
    grid-row: 2 / span 1;
    grid-column: 14 / span 1;
  }
.sm-box--15 {
    grid-row: 2 / span 1;
    grid-column: 15 / span 1;
  }
 .sm-box--16 {
    grid-row: 2 / span 1;
    grid-column: 16 / span 1;
  }
.sm-box--17 {
    grid-row: 2 / span 1;
    grid-column: 17 / span 1;
  }
.sm-box--18 {
    grid-row: 2 / span 1;
    grid-column: 18 / span 1;
  }
.sm-box--19 {
    grid-row: 2 / span 1;
    grid-column: 19 / span 1;
  }
.sm-box--20 {
    grid-row: 2 / span 1;
    grid-column: 20 / span 1;
  }
 .sm-box--21 {
    grid-row: 2 / span 1;
    grid-column: 21 / span 1;
  }
.sm-box--22 {
    grid-row: 2 / span 1;
    grid-column: 22 / span 1;
  }
.sm-box--23 {
    grid-row: 2 / span 1;
    grid-column: 23 / span 1;
  }
.sm-box--24 {
    grid-row: 2 / span 1;
    grid-column: 24 / span 1;
  }
  .sm-box--25 {
    grid-row: 3 / span 1;
    grid-column: 1 / span 1;
  }
.sm-box--26 {
    grid-row: 3 / span 1;
    grid-column: 2 / span 1;
  }
.sm-box--27 {
    grid-row: 3 / span 1;
    grid-column: 3 / span 1;
  }
.sm-box--28 {
    grid-row: 3 / span 1;
    grid-column: 4 / span 1;
  }
.sm-box--29 {
    grid-row: 3 / span 1;
    grid-column: 5 / span 1;
  }
.sm-box--30 {
    grid-row: 3 / span 1;
    grid-column: 6 / span 1;
  }
.sm-box--31 {
    grid-row: 3 / span 1;
    grid-column: 7 / span 1;
  }
.sm-box--32 {
    grid-row: 3 / span 1;
    grid-column: 8 / span 1;
  }
.sm-box--33 {
    grid-row: 3 / span 1;
    grid-column: 9 / span 1;
  }
.sm-box--34 {
    grid-row: 3 / span 1;
    grid-column: 10 / span 1;
  }
.sm-box--35 {
    grid-row: 3 / span 1;
    grid-column: 11 / span 1;
  }
.sm-box--36 {
    grid-row: 3 / span 1;
    grid-column: 12 / span 1;
  }
.sm-box--37 {
    grid-row: 3 / span 1;
    grid-column: 13 / span 1;
  }
.sm-box--38 {
    grid-row: 3 / span 1;
    grid-column: 14 / span 1;
  }
.sm-box--39 {
    grid-row: 3 / span 1;
    grid-column: 15 / span 1;
  }
.sm-box--40 {
    grid-row: 3 / span 1;
    grid-column: 16 / span 1;
  }
 .sm-box--41 {
    grid-row: 3 / span 1;
    grid-column: 17 / span 1;
  }
.sm-box--42 {
    grid-row: 3 / span 1;
    grid-column: 18 / span 1;
  }
.sm-box--43 {
    grid-row: 3 / span 1;
    grid-column: 19 / span 1;
  }
.sm-box--44 {
    grid-row: 3 / span 1;
    grid-column: 20 / span 1;
  }
.sm-box--45 {
    grid-row: 3 / span 1;
    grid-column: 21 / span 1;
  }
 .sm-box--46 {
    grid-row: 3 / span 1;
    grid-column: 22 / span 1;
  }
.sm-box--47 {
    grid-row: 3 / span 1;
    grid-column: 23 / span 1;
  }
.sm-box--48 {
    grid-row: 3 / span 1;
    grid-column: 24 / span 1;
  }
</style>

<div class="grid-container">
  <header>$fileName</header>
  <div class="sm-box sm-box--1" style="background-color:$port1;$splitport1">1</div>
  <div class="sm-box sm-box--2" style="background-color:$port2;$splitport2">2</div>
  <div class="sm-box sm-box--3" style="background-color:$port3;$splitport3">3</div>
  <div class="sm-box sm-box--4" style="background-color:$port4;$splitport4">4</div>
  <div class="sm-box sm-box--5" style="background-color:$port5;$splitport5">5</div>
  <div class="sm-box sm-box--6" style="background-color:$port6;$splitport6">6</div>
  <div class="sm-box sm-box--7" style="background-color:$port7;$splitport7">7</div>
  <div class="sm-box sm-box--8" style="background-color:$port8;$splitport8">8</div>
  <div class="sm-box sm-box--9" style="background-color:$port9;$splitport9">9</div>
  <div class="sm-box sm-box--10" style="background-color:$port10;$splitport10">10</div>
  <div class="sm-box sm-box--11" style="background-color:$port11;$splitport11">11</div>
  <div class="sm-box sm-box--12" style="background-color:$port12;$splitport12">12</div>
  <div class="sm-box sm-box--13" style="background-color:$port13;$splitport13">13</div>
  <div class="sm-box sm-box--14" style="background-color:$port14;$splitport14">14</div>
  <div class="sm-box sm-box--15" style="background-color:$port15;$splitport15">15</div>
  <div class="sm-box sm-box--16" style="background-color:$port16;$splitport16">16</div>
  <div class="sm-box sm-box--17" style="background-color:$port17;$splitport17">17</div>
  <div class="sm-box sm-box--18" style="background-color:$port18;$splitport18">18</div>
  <div class="sm-box sm-box--19" style="background-color:$port19;$splitport19">19</div>
  <div class="sm-box sm-box--20" style="background-color:$port20;$splitport20">20</div>
  <div class="sm-box sm-box--21" style="background-color:$port21;$splitport21">21</div>
  <div class="sm-box sm-box--22" style="background-color:$port22;$splitport22">22</div>
  <div class="sm-box sm-box--23" style="background-color:$port23;$splitport23">23</div>
  <div class="sm-box sm-box--24" style="background-color:$port24;$splitport24">24</div>
  <div class="sm-box sm-box--25" style="background-color:$port25;$splitport25">25</div>
  <div class="sm-box sm-box--26" style="background-color:$port26;$splitport26">26</div>
  <div class="sm-box sm-box--27" style="background-color:$port27;$splitport27">27</div>
  <div class="sm-box sm-box--28" style="background-color:$port28;$splitport28">28</div>
  <div class="sm-box sm-box--29" style="background-color:$port29;$splitport29">29</div>
  <div class="sm-box sm-box--30" style="background-color:$port30;$splitport30">30</div>
  <div class="sm-box sm-box--31" style="background-color:$port31;$splitport31">31</div>
  <div class="sm-box sm-box--32" style="background-color:$port32;$splitport32">32</div>
  <div class="sm-box sm-box--33" style="background-color:$port33;$splitport33">33</div>
  <div class="sm-box sm-box--34" style="background-color:$port34;$splitport34">34</div>
  <div class="sm-box sm-box--35" style="background-color:$port35;$splitport35">35</div>
  <div class="sm-box sm-box--36" style="background-color:$port36;$splitport36">36</div>
  <div class="sm-box sm-box--37" style="background-color:$port37;$splitport37">37</div>
  <div class="sm-box sm-box--38" style="background-color:$port38;$splitport38">38</div>
  <div class="sm-box sm-box--39" style="background-color:$port39;$splitport39">39</div>
  <div class="sm-box sm-box--40" style="background-color:$port40;$splitport40">40</div>
  <div class="sm-box sm-box--41" style="background-color:$port41;$splitport41">41</div>
  <div class="sm-box sm-box--42" style="background-color:$port42;$splitport42">42</div>
  <div class="sm-box sm-box--43" style="background-color:$port43;$splitport43">43</div>
  <div class="sm-box sm-box--44" style="background-color:$port44;$splitport44">44</div>
  <div class="sm-box sm-box--45" style="background-color:$port45;$splitport45">45</div>
  <div class="sm-box sm-box--46" style="background-color:$port46;$splitport46">46</div>
  <div class="sm-box sm-box--47" style="background-color:$port47;$splitport47">47</div>
  <div class="sm-box sm-box--48" style="background-color:$port48;$splitport48">48</div>
  <div class="g-box g-box--1" style="background-color:$gport1;">G1</div>
  <div class="g-box g-box--2" style="background-color:$gport2;">G2</div>
  <div class="g-box g-box--3" style="background-color:$gport3;">G3</div>
  <div class="g-box g-box--4" style="background-color:$gport4;">G4</div>
  <div class="key-box key-box--1" style="background-color:#7FFF00;">Staff</div>
  <div class="key-box key-box--2" style="background-color:#DC143C;">Public</div>
  <div class="key-box key-box--3" style="background-color:#0000FF;">Voice</div>
  <div class="key-box key-box--4" style="background-color:#9932CC;">WIFI</div>
  <div class="key-box key-box--5" style="background-color:#E67E22;">RFID</div>
  <div class="key-box key-box--6" style="background-color:#FF00FF;">Trunk</div>
  <div class="key-box key-box--7" style="background-color:#616A6B;">Other</div>
</div>
"@

$html | set-content -path "c:\scripts\switch\$fileName.html"

}

#delete old html files
Get-ChildItem -Path "c:\scripts\switch\" | Remove-Item

#extract files from JAR
$files = Get-ChildItem -Path $source | Where-Object {$_.LastWriteTime -ge (get-date -hour 0 -minute 0)}

ForEach ($f in $files) {
    [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
    [System.IO.Compression.ZipFile]::ExtractToDirectory($f.FullName, $target)

    # DO STUFF WITH FILES
    $configFile = Get-ChildItem -Path $target | Where-Object {$_.Name -match $pattern}

    $content = Get-Content $configFile.FullName
    
    for ($i=0; $i -le ($content.Length -1); $i += 1) {
        if ($record) { Add-ToExistingObject }
        switch -Wildcard ($content[$i]) {
            "interface g*" { Add-ToNewObject -type "g"; $record = $true  }
            "interface f*" { Add-ToNewObject -type "f"; $record = $true  }
            "!" { if ($record) {Close-Object}; $record = $false }
        }
        
    }
    
    if ($array.count -gt 40) {
        48Port-HTML -swName $configFile.BaseName
    } else {
        24Port-HTML -swName $configFile.BaseName
    }

    # DELETE FILES BEFORE MORE COME OUT
    Get-ChildItem -Path $target | Remove-Item
    # RESET VARIABLES
    $record = $false
    $global:array = @()
    $global:interface = $null
}
