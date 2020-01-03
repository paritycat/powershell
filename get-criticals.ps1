#code : rn1737/rni-a62162 2016-2020
clear-host
write-host "Get-Critical : Analyse des erreurs critiques du système d'exploitation"
write-host "======================================================================"
write-host ""
try
{
$criticals = Get-WinEvent system | Where-Object {$_.Level -eq 1 -and $_.ProviderName -eq "Microsoft-Windows-Kernel-Power"} | Select-Object timecreated, message
}
catch
{
write-host "Impossible de récupérer les évènements d'arrêt critiques !" -ForegroundColor Red
break
}

$nbcriticals = $criticals.Count 

# il faut a minima deux évènements pour faire les stats

if ($nbcriticals -eq 1)

{

Write-Host "WARNING : un seul évènement d'arrêt critique enregistré. Aucune statistique avancée ne sera fournie. Voici le détail ci dessous : " -ForegroundColor Yellow
$criticals | ft
write-host
$dummy = Read-Host "Appuyez sur une touche pour fermer l'utilitaire..."

}


if ($nbcriticals -eq 0)

{

Write-Host "OK : aucun évènement d'arrêt critique enregistré. Soit votre PC semble sain, soit aucun enregistrement n'est disponible dans le journal Système." -ForegroundColor Yellow
$criticals | ft
write-host
$dummy = Read-Host "Appuyez sur une touche pour fermer l'utilitaire..."

}

if ($nbcriticals -ge 2)

{

    try{

        $compteur_defaillances = $criticals.Count # compteur
        $firstcritical = ($criticals | Select-Object -First 1).timecreated # plus récent
        $lastcritical = ($criticals | Select-Object -Last 1).timecreated # plus ancien (donc le premier)
        $daysbetween = (get-date) - ($firstcritical) #denier arrêt il y a X jours

        $totalinterval = $firstcritical - $lastcritical #intervalle d'analyse entre les deux arrêts (le + ancien / le + récent)
        $moyenne_semaine = [math]::Round((($nbcriticals / $totalinterval.Days) * 7),2) # arrêts par semaine
        $moyenne_mois = [math]::Round((($nbcriticals / $totalinterval.days) * 30.5) , 2) # mois :)
        $moyenne_an = [math]::Round((($nbcriticals / $totalinterval.days) * 365) , 2) # an :)
        $ratioplantage = [math]::Round((365.25 / $moyenne_an),0) #ratio de plantage

        write-host "Dernier arrêt critique du système   :" $firstcritical
        write-host "Premier arrêt critique du système   :" $lastcritical
        write-host "Intervalle depuis le dernier arrêt  :" $daysbetween.Days  "jour(s) et " $daysbetween.hours "heure(s)"
        write-host "Nombre d'arrêts critiques monitorés :" $nbcriticals
        write-host "Intervalle entre 1er & dernier arrêt:" $totalinterval.Days "jours(s)"
        write-host "Moyenne sur l'intervalle 1er/dernier:" $moyenne_semaine "par semaine," $moyenne_mois "par mois ou" $moyenne_an "par an"
        write-host "                                      Le PC plante donc en moyenne une fois tous les" $ratioplantage "jours"
        write-host
        $dummy = Read-Host "Appuyez sur une touche pour fermer l'utilitaire..."

    }

    catch

    {

    write-host "Impossible de faire les statistiques sur les évènements d'arrêt critiques !" -ForegroundColor Red
    break

    }
}





