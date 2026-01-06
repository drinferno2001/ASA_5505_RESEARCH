###########################
# VARIABLE INITIALIZATION #
###########################

#############################################
# ADDRESS RANGES TO PULL (SPECIFIED AS HEX) #
#############################################

$addrRanges = @(
    @{
        Start = "B15F6BB4";
        End = "B15F6DB4";
    }
)

########################
# SERIAL PORT SETTINGS #
########################

$portName = "COM6";
$baudRate = 9600;
$parity = [System.IO.Ports.Parity]::None;
$dataBits = 8;
$stopBits = [System.IO.Ports.StopBits]::One;
$serialPort = New-Object System.IO.Ports.SerialPort $portName, $baudRate, $parity, $dataBits, $stopBits;
$serialPort.ReadTimeout = 5000;
$serialPort.WriteTimeout = 5000;

###################
# SCRIPT SETTINGS #
###################

# Executing Interval (In Milliseconds)
# Description: How often do we execute a command to pull data?
$executionInterval = 5000

##########
# SCRIPT #
##########

Write-Host "";

####################
# OPEN SERIAL PORT #
####################

try{
    $serialPort.Open()
    Write-Host "LOG: SERIAL PORT [$($portName)] OPENED SUCCESSFULLY";
}catch{
    Write-Error "FAILED TO OPEN SERIAL PORT: $_";
    exit;
}

##################################################
# LOOP THROUGH AND PULL SPECIFIED ADDRESS RANGES #
##################################################

$exportStream = $null;
$hexDumpPath = $null;

try{
    foreach($range in $addrRanges){

        # Define numeric (and printable) address range values
        $rangeInfo = @{
            Start = @{
                Decimal = [Convert]::ToInt64($range.Start, 16);
                Hex = $range.Start;
            }
            End = @{
                Decimal = [Convert]::ToInt64($range.End, 16);
                Hex = $range.End;
            }
        }

        # Define date formats
        $processStartDate = (Get-Date).toString("MM/dd/yyyy hh:mm tt");
        $exportStartDate = (Get-Date).toString("MM-dd-yyyy_hh_mm_tt");

        Write-Host "LOG: DUMPING ADDRESS RANGE [($($rangeInfo.Start.Hex)) - ($($rangeInfo.End.Hex))] ON [$($processStartDate)]";

        ###########################################
        # OPEN A CONNECTION TO BINARY EXPORT FILE #
        ###########################################

        $hexDumpPath = "$($PSScriptRoot)\$($rangeInfo.Start.Hex)_-_$($rangeInfo.End.Hex)_STARTED_ON_$($exportStartDate).bin";

        try{
            $exportStream = [System.IO.File]::Open($hexDumpPath, [System.IO.FileMode]::Append)
            Write-Host "LOG: ASSOCIATED BINARY EXPORT FILE [$($hexDumpPath)] OPENED/CREATED SUCCESSFULLY";

        }catch{
            throw "ERROR CREATING/OPENING BINARY EXPORT FILE AT [$($hexDumpPath)]: $_";
        }

        # Identify start address as the initial current address when processing
        $rangeInfo.Current = @{
            Decimal = $rangeInfo.Start.Decimal
            Hex = $rangeInfo.Start.Hex
        }

        try{
            # Loop through contents of addresses in each range
            while($rangeInfo.Current.Decimal -lt $rangeInfo.End.Decimal){

                $command = "D %$($rangeInfo.Current.Hex)";
                $addressContents = $null;

                try{
                    Write-Host "LOG: EXECUTING COMMAND: [$($command)]";

                    # Send command to output address
                    $serialPort.Write("");
                    $serialPort.Write("$($command)`r");

                    # Give the device time to respond
                    Start-Sleep -Milliseconds $executionInterval;

                    # Read string response
                    $addressContents = $serialPort.ReadExisting();

                    if(-Not($addressContents)){
                        throw "NO RESPONSE RECEIVED";
                    }
                }catch{
                    throw "ERROR EXECUTING MEMORY PULL COMMAND [$($command)]: $_";
                }

                try{

                    #########################################################
                    # PARSE RESPONSE AND PULL HEX CHARACTERS FROM EACH LINE #
                    #########################################################
                
                    # Break response into list of lines
                    $lines = @($addressContents -split '\r?\n');

                    $lines = @($lines | Where-Object{ 
                        $_ -match '^0{4}\:[0-9,A-Z,a-z]{8}\ (\ [0-9,A-Z,a-z]{2}){8}\:([0-9,A-Z,a-z]{2}\ ){8}.{16}$'
                    })

                    # Parse lines to remove un-necessary content
                    $lines | ForEach-Object{
                        
                        # Only pull hex digits (and remove ':' in middle)
                        Write-Verbose -verbose $_;
                        $currentLine = $_.substring(15, 47);
                        $currentLine = $currentLine -replace ":"," "
                        
                        # Get the contents of the current line
                        $lineContents = @($currentLine -split "\s+");

                        # NOTE: AT THIS POINT, THERE SHOULD BE 2-HEX CHARACTER PATTERNS DELIMITED BY SPACES

                        # Convert to binary by prefixing '0X' to each 2-hex-digit string
                        # (and then casting to a byte array)

                        [byte[]]$bytes = ($lineContents | ForEach-Object{$_ -replace '^', '0X'})

                        # Write the new bytes
                        $exportStream.Write($bytes, 0, $bytes.Length)
                    }
                }catch{
                    throw "ERROR CREATING/APPENDING TO BINARY EXPORT: $($_)"
                }

                # Increment in values of 128

                $rangeInfo.Current.Decimal += 128;
                $rangeInfo.Current.Hex = "{0:X8}" -f $rangeInfo.Current.Decimal;
            }
        }catch{
            throw "ERROR EXECUTING DUMP OF CURRENT ADDRESS RANGE: $_";
        }

        $endDate = (Get-Date).toString("MM/dd/yyyy hh:mm tt");
        Write-Host "LOG: FINISHED DUMP OF ADDRESS RANGE [($($rangeInfo.Start.Hex)) - ($($rangeInfo.End.Hex))] ON [$($endDate)]";
    }
}catch{
    Write-Error $_
}finally{

    #################################################################
    # CLOSE ADDRESS RANGE BINARY EXPORT FILE HANDLE (IF APPLICABLE) #
    #################################################################

    if ($exportStream) {
        $exportStream.Dispose();
        $exportStream = $null;

        Write-Host "LOG: BINARY EXPORT FILE [$($hexDumpPath)] HAS BEEN CLOSED";
    }

    #####################
    # CLOSE SERIAL PORT #
    #####################

    $serialPort.Close()
    Write-Host "LOG: SERIAL PORT [$($portName)] HAS BEEN CLOSED";
}

Write-Host "";