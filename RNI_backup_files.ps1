#                  ````............``````                        ````.............```                
#             ``.............................`               `..........................``           
#         ``..................................            `.................................`        
#       `.............```         `````.......          `...........```         ``............`      
#     `...........`                        ```        `..........`                  `...........`    
#   `..........`                                     ..........`                       ..........`   
#  `..........                                      ..........                          `.........`  
# `.........`                                      ..........                            `.........` 
# ..........                                      `.........`               `---------.   .......... 
#`.........                                       ..........              `:++++++++/.    `.........`
#..........                                       ..........            `-++++++++/.      `.........`
#..........                                       ..........          `:++++++++:`         .........`
#..........                                       ..........        .:+++++++:.           `.........`
#`.........`                                      `.........     `-/+++++/:.              `......... 
# ..........`                                      .........`   .------.`                 .......... 
#  ..........`                                     `.........`                           ..........  
#   ...........`                                    `.........`                        `..........   
#    `...........`                                   `..........`                    `..........`    
#      .............```                ```....         `...........`              ``...........`     
#        `....................................           `..............``````...............`       
#           `.................................             ``.............................`          
#               ``........................````                 ``....................``              
#                      `````````````                                   `````````                     
#                                              `........``                                           
#.::::::                        -:::::`   `.:/++++++++++++/:.`                 `.--::::::::::------- 
# -+++++/`                     :++++/.  -/++++++/::--::/++++++:`           `-:/++++++++++++++++++++/ 
#  -+++++/`                   :++++/` ./+++++:.         `./+++++-        `:++++++/:-..````...-+++++/ 
#   .+++++/`                `/++++/` :+++++/`              :+++++.     `:+++++/-`            `+++++/ 
#    .+++++/`              `/++++/` :+++++/                `+++++/    ./+++++-               `+++++/ 
#     `/+++++.            `/++++:  .++++++-................./+++++.  .++++++.                `+++++/ 
#      `/+++++.          .+++++:   :++++++++++++++++++++++++++++++-  /+++++-                 `+++++/ 
#       `/+++++-        .+++++-    :+++++/........................` .++++++`                 `+++++/ 
#         :+++++-      -+++++-     -+++++/                          -+++++/                  .+++++/ 
#          :+++++:    :+++++.       /+++++-                         -++++++`                -++++++/ 
#           :+++++:  :++++/.        `/+++++:`                       `++++++:              -/+++++++/ 
#            -+++++//++++/`           -++++++:.              `.-::   -++++++:`         .:++++:+++++/ 
#             -+++++++++/`             `-/++++++/:-------::/+++++/    -+++++++:-....-:+++++:` /++++/ 
#              .+++++++/`                 .-:/+++++++++++++++/::-`     `-/++++++++++++++:.    :+++++.
#               `------                       ``.-------..``              `-://////::-`       `:::::-
#
# Sauvegarde d'un répertoire utilisateur avec compression ZIP et envoi de mail
# Outil de backup rudimentaire :)
# RN 08/2018
# V.0.1 Initiale le 30/07/2018
# V 0.2 : Ajout : suppression du répertoire source après création du ZIP
#         Modifications pour utiliser le script partout (+ de variables,[INFO]`tde données en "dur")
# V 0.3 : Correction de bugs divers (création du répertoire)
#         Ajout des options de personnalisation du comportement du script
#         Mise au propre du code
# V 1.0 : Janvier 2019 --> réécriture partielle (compression ZIP etc.)
####################################################################################################


#region fonction (Envoi HTML)

function envoiHTML
{
<#  .SYNOPSIS
    Envoie un mail au format HTML

    .DESCRIPTION
    Envoie un message depuis un serveur sans authentification au format HTML
    Permet l'ajout de PJ (ou non :)

    .EXAMPLE
    envoiHTML -server "monserveur.domaine.extension" -to "destinataire@domaine.ext" -from "expediteur@domaine.ext" -subject "Objet du mail" -attachment "chemin\fichier" -bodyhtm "Message au format HTM<br>Ligne<hr>etc."

    .PARAMETER Server Contient le FQDN (nom de domaine pleinement qualifié) du serveur de mail, sans authentification
    .PARAMETER Body Contient le corps du mail au format HTML, on peut y ajouter les balises habituelles (img, hr, br etc.)
    .PARAMETER From Contient le mail de l'expéditeur 
    .PARAMETER To Contient le destinataire du mail
    .PARAMETER Subject Contient l'objet du mail
    .PARAMETER Attachment Contient un fichier en pièce jointe (facultatif)...le fichier doit exister sans quoi :)))
#>  

param ([string]$server , [string]$to, [string]$from, [string]$subject, [string]$attachment, [string]$bodyhtm)

$encoding = New-Object System.Text.utf8encoding


if ($attachment)

    {
        try
            {

                send-MailMessage -SmtpServer $server -To $to -From $from -Subject $subject -Body $bodyhtm -BodyAsHtml -Attachments $attachment -Encoding "UTF8"
            
            }
        
        catch
            
            {

                Write-host "Erreur d'envoi du message, this is some serious sh*t :) " $_.exception
            
            }
    }

else

    {

     try
            {
            
            send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml 
            
            }
        
        catch
            
            {
            
            Write-host "Erreur d'envoi du message, this is some serious sh*t :) "
            write-host $_.errormessage
            
            }

    }

}

#endregion

#region raz/prérequis&options utilisateurs

clear-host

#Ajout de l'assembly de compression, mise  à 0 erreurs + récupération date à un format acceptable par un nom de fichier, pour le zip
write-host "[INFO]`tDémarrage du script de sauvegarde en cours..."
#add-type -assembly "system.io.compression.filesystem" 


#RAZ compteur erreur
$errorcount = 0 
$warningcount = 0

#Mise en place de la date
$ajd = Get-Date -UFormat "%d-%m-%Y" #récupération de la date pour créer le fichier

write-host "[INFO]`tImport des informations de backup"
##################################################################EDITEZ LE FICHIER DEPUIS CETTE LIGNE #################################
# Toute les lignes avec un [W] sont des options éditables :)

#Fichiers à copier (source et destination)
#Les répertoires doivent exister !!

$source = 'C:\Users\A62162\Documents\' #Source des fichiers à copier [W]
$dossieracreer = "automatique" #Nom du répertoire de destination à créer pour la création du ZIP [W]
$destination = '\\ads01.priv\Documents\Clients\SAR\A62162\' + $dossieracreer #Destination des fichiers [W]
write-host "[INFO]`tDossier source  : " $source
write-host "[INFO]`tDossier à créer : " $dossieracreer
write-host "[INFO]`tDestination     : " $destination

#Construction du fichier ZIP : source des fichiers à zipper et nommage

$sourcezip = $destination #Emplacement où sont les fichiers pour créer le ZIP de backup (généralement la destination des fichiers)  [W]
$fichier = "mydocs-$ajd.zip" #convention de nommage du fichier ZIP  [W]
$backup = $destination #emplacement où sera mis le ZIP (répertoire). Si la destination n'a pas beaucoup d'espace, le placer ailleurs :)  [W]
$destinationzip = $backup + $fichier #création chemin + nom du fichier ZIP

write-host "[INFO]`tFichier ZIP     : " $fichier

#Mail d'envoi
$mail = "rnassiri.ext@gmf.fr"  #Mail du destinataire (est également celui de l'expéditeur) [W]
$mailsubject = "Rapport de sauvegarde pour " + $mail #Objet du mail [W]
$mailserver = "cassar01.ads01.priv" #Serveur utilisé pour l'envoi (sans authentification) [W]
write-host "[INFO]`tDestinataire    : " $mail
write-host "[INFO]`tServeur SMTP    : " $mailserver

#Emplacement du log

$backuplog = "\\ads01.priv\Documents\Clients\SAR\A62162\backuplog-$ajd.txt" #nom du fichier LOG [W]
write-host "[INFO]`tFichier de log  : " $backuplog

#Option : backup Outlook
#Si cette variable est à $true, le script fermera automatiquement Outlook afin de sauvegarder d'éventuels fichiers PST
#qui seraient présents dans le répertoire source. La sauvegarde des PST ouverts dans MSO n'est pas possible, il faut que 
#Outlook soit fermé

$outlookpst = $false # [W]

#Option : suppression finale du répertoire de backup
#Cela explique pourquoi on copie d'abord les fichiers vers une destination autre 
#Si mis à $true, supprimera le répertoire de destination pour ne créer que le ZIP

$deletedestination = $true # [W]

#Option : créer le ZIP
#Si cette option est à $true, créé un ZIP pour compresser tout cela.
#Conseil : désactiver ($false) si votre répertoire source est trop gros ou si la destination ne pourra pas accueillir ZIP et fichiers bruts
#Si $deletedestination est à $true et cette option à $false, votre copie n'aura servi à rien

$createzip = $true # [W]

#Option : envoyer un mail
#Si cette option est à $true, envoie un mail de récap avec les informations présentes plus haut
#sinon, pas de mail

$sendmail = $true # [W]

write-host "[INFO]`tFermer Outlook  : " $outlookpst
write-host "[INFO]`tSupp. source ZIP: " $deletedestination
write-host "[INFO]`tCréer un ZIP    : " $createzip
write-host "[INFO]`tEnvoyer un mail : " $sendmail

###################################################################FIN DE LA ZONE D'EDITION DES FICHIERS ##############################

#endregion

#region code



        
        

        $datelog = Get-Date -UFormat "%H:%M:%S"
        $outfile = $datelog + " [INFO]`t" + "Démarrage de la sauvegarde" | out-file $backuplog #Créer le fichier de log...
        
        # v1. Paramètres du backup dans le LOG :

        $outfile = $datelog +  " [INFO]`t##Informations sur le backup##" | out-file $backuplog
        $outfile = $datelog +  " [INFO]`tDossier source  : " + $source | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tDossier à créer : " +$dossieracreer | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tDestination     : " +$destination | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tFichier ZIP     : " +$fichier | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tDestinataire    : " +$mail | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tServeur SMTP    : " +$mailserver | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tFichier de log  : " +$backuplog | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tFermer Outlook  : " +$outlookpst | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tSupp. source ZIP: " +$deletedestination | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tCréer un ZIP    : " +$createzip | out-file $backuplog -Append
        $outfile = $datelog +  " [INFO]`tEnvoyer un mail : " +$sendmail | out-file $backuplog -Append


        #Ferme Outlook si des PSTs sont présents dans le répertoire à copier, la copie d'un PST n'étant pas possible outlook ouvert...

        if ($outlookpst -eq $true)

            {

                try
        
                    {
                        write-host "[INFO]`tFermeture de Outlook..."
                        Get-Process outlook| Foreach-Object { $_.CloseMainWindow() } -ErrorAction SilentlyContinue | Out-Null
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [INFO]`t" + "Outlook fermé..." | out-file $backuplog -Append
        
                    }
        
                catch
        
                    {
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + "[WARN]`t" + "Outlook est déjà fermé..." | out-file $backuplog -Append
                        $errorcount++
        
                    }
        
            }

        else

            {
                
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + "Pas de détection Outlook/PST demandée !" | out-file $backuplog -Append

            }

        #Création du répertoire s'il n'existe pas 
        try

            {
                New-Item -Name $destination -ItemType directory -ErrorAction SilentlyContinue
                write-host "[INFO]`tDossier" $destination "créé"
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + "Dossier de destination " + $destination + " créé." | out-file $backuplog -Append
            }

        catch
        
            {
                write-host "[WARN]`tImpossible de créer le dossier, peut-être existe-t-il déjà ?"
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + "[ERREUR]`t" + "Impossible de créer le dossier, peut-être existe-t-il déjà." | out-file $backuplog -Append
                $warningcount++
            }

        # v1.0 analyse des fichiers :
        $Files=@()
        $SumMB=0
        $SumItems=0
        $SumCount=0
        $colItems=0
        $colItems = (Get-ChildItem $source -recurse | Where-Object {$_.mode -notmatch "h"} | Measure-Object -property length -sum) 
        $FilesCount += Get-ChildItem $source -Recurse | Where-Object {$_.mode -notmatch "h"}  
        $SumMB+=$colItems.Sum.ToString()
        $SumItems+=$colItems.Count
        $TotalMB="{0:N2}" -f ($SumMB / 1MB) + " Mo de fichiers"
        write-host "[INFO]`tIl y a $SumItems fichiers pour  $TotalMB à copier."
        $datelog = Get-Date -UFormat "%H:%M:%S"
        $outfile = $datelog + " [INFO]`tDécompte du volume à copier : " + $SumItems + " fichiers et " + $TotalMB + " à copier." | out-file $backuplog -Append

        #Copie des fichiers
        
        
        try

            {
                
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + "Copie des fichiers démarrée" | out-file $backuplog -Append
                write-host "[INFO]`tCopie des fichiers de " $source "vers la destination " $destination " ..."
                # v1.0 analyse du temps de copie
                $timetaken = Measure-Command {
                copy-item -Path $source -Destination $destination -Recurse -Force -erroraction silentlycontinue -Verbose 
                }
                write-host "[INFO]`tDurée écoulée : " $timetaken.Hours " h " $timetaken.Minutes " min. " $timetaken.Seconds " s."
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`tDurée écoulée : " + $timetaken.Hours + " h " + $timetaken.Minutes + " min. "+ $timetaken.Seconds + " s." | out-file $backuplog -Append

                $bytescopied = $SumMB
                $totalsecond = $timetaken.TotalSeconds
                $mbytesperseconds = ($SumMB/1024/1024) / $totalsecond 
                write-host "[INFO]`tDébit (Mo/s) : " $mbytesperseconds
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`tDébit (Mo/s) : " + $mbytesperseconds | out-file $backuplog -Append

                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + "Fin de la copie des fichiers" | out-file $backuplog -Append
                Write-Host "[INFO]`tCopie terminée"
        
            }

        catch

            {
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + " Erreur dans la copie des fichiers depuis :" + $source + " vers : " + $destination | out-file $backuplog -Append
                $errorcount++
            }
                  
        
        #Suppression du ZIP s'il existe déjà :)
             
        
        If(Test-path $destinationzip) 

            {
                try
                    {
                        write-host "[INFO]`tVérification présence ZIP"
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [INFO]`t" + "Vérification ZIP" | out-file $backuplog -Append
                        Remove-item $destinationzip -ErrorAction SilentlyContinue
                        write-host "[INFO]`tUn fichier ZIP existe déjà, suppression..."
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [WARN]`t" + "Déjà existant, supprimé" | out-file $backuplog -Append
                    }
                
                catch
                
                    {
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + "[ERREUR]`t" + "Erreur dans la suppression du ZIP : " + $destination | out-file $backuplog -Append
                        $errorcount++
                    }
            
            }

        #Compression en ZIP :)
        
        
            if ($createzip -eq $true)
                
                {
                
                try
                
                    {
                        write-host "[INFO]`tCompression en cours..."
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [INFO]`t" + "Démarrage compression" | out-file $backuplog -Append
                        Compress-Archive -Path $sourcezip -DestinationPath $destinationzip -CompressionLevel Optimal -Force -Verbose
                        #[io.compression.zipfile]::CreateFromDirectory($sourcezip, $destinationzip) 
                        write-host "[INFO]`tCompression terminée"
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [INFO]`t" + "Fin de la compression dans le fichier : " + $destinationzip | out-file $backuplog -Append

                    }
                
                catch
                
                    {
                        $datelog = Get-Date -UFormat "%H:%M:%S"
                        $outfile = $datelog + " [INFO]`t" + "Erreur dans la compression ZIP vers : " + $destinationzip | out-file $backuplog -Append
                        write-host "[ERROR]`tCompression terminée"
                        $errorcount++
                    }
                }
            
            else
            
                {
                    $datelog = Get-Date -UFormat "%H:%M:%S"
                    $outfile = $datelog + " [INFO]`t" + "Pas de création de ZIP à la demande de l'utilisateur " | out-file $backuplog -Append
                    
                }
        
        #Suppression du répertoire $sourcezip 
            
            if ($deletedestination -eq $true)

                {

                    try
                        {
                            $datelog = Get-Date -UFormat "%H:%M:%S"
                            $outfile = $datelog + " [INFO]`t" + "Suppression du répertoire : " + $destination | out-file $backuplog -Append
                            write-host "[INFO]`tSuppression du répertoire de source ZIP"
                            Remove-Item $destination -Recurse -Force
                            write-host "[INFO]`tSuppression terminée !"
                        }
                
                    catch
                
                        {
                            $datelog = Get-Date -UFormat "%H:%M:%S"
                            $outfile = $datelog + " [INFO]`t" + "Erreur dans la suppression du répertoire : " + $destination | out-file $backuplog -Append
                            write-host "[ERROR]`tErreur de suppression du répertoire de source ZIP"
                            $errorcount++
                        }

                }
            
            else
            
                {
                    $datelog = Get-Date -UFormat "%H:%M:%S"
                    $outfile = $datelog + " [INFO]`t" + "Le répertoire de destination ne sera pas supprimé. Conserve : " + $destination | out-file $backuplog -Append
                    write-host "[INFO]`tSuppression du répertoire de destination non demandé."
                    $errorcount++
                }

        
        #Envoi du mail (notez que les deux dernières lignes ne sont pas incluses dans le fichier LOG, déjà envoyé)
        

        if ($sendmail -eq $true)

        {
            try

                {
                    write-host "[INFO]`tEnvoi du mail en cours..."
                    $datelog = Get-Date -UFormat "%H:%M:%S"
                    $outfile = $datelog + " [INFO]`t" + "Envoi du mail" | out-file $backuplog -Append
                    $bodyhtm = "Sauvegarde effectuée à vérifier SVP<hr><br>La sauvegarde est terminée.<br>" + [string]$errorcount  + " erreurs[INFO]`t" + [string]$warningcount  + " warnings"
                    envoiHTML -server $mailserver -to $mail -from $mail -subject $mailsubject -bodyhtm $bodyhtm -attachment $backuplog
                    write-host "[INFO]`tEnvoi terminé"
                    $datelog = Get-Date -UFormat "%H:%M:%S"
                    $outfile = $datelog + " [INFO]`t" + "FIN : Mail envoyé (ne sera pas présent dans le fichier joint au mail)" | out-file $backuplog -Append
                    write-host "[INFO]`t" $errorcount "erreur(s) -" $warningcount "warning(s)"
                    start-sleep -Seconds 2
                }

            catch

                {
                    $datelog = Get-Date -UFormat "%H:%M:%S"
                    $outfile = $datelog + " [INFO]`t" + "Mail non envoyé : erreur durant l'envoi" | out-file $backuplog -Append
                    write-host "[INFO]`t" $errorcount "erreur(s) -" $warningcount "warning(s)"
                }

        }

        else

            {
                $datelog = Get-Date -UFormat "%H:%M:%S"
                $outfile = $datelog + " [INFO]`t" + "Pas d'envoi de mail (à la demande de l'utilisateur)" | out-file $backuplog -Append
                write-host "[INFO]`t pas d'envoi de mail demandé"
            
            }

#endregion


#FIN des opérations