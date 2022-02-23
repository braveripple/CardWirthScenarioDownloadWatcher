<###############################################################
CardWirthScenarioDownloadWatcher
###############################################################>
#Requires -Version 3
#Requires -Modules @{ ModuleName="CardWirthScenarioSummaryReader"; ModuleVersion="1.0.0" }

using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms;

# 設定してください
$WATCH_DIRECTORY_PATH = "$HOME/Downloads"
$CARDWIRTH_SCENARIO_DIRECTORY_PATH = "$HOME/CardWirthPy/Scenario/新着シナリオ/"
$CARDWIRTHNEXT_SCENARIO_DIRECTORY_PATH = "$HOME/CardWirthNext/Scenario/新着シナリオ/"

$TIMER_INTERVAL = 3 * 1000 # timer_function実行間隔(ミリ秒)

$MUTEX_NAME = '0b72e703-1de1-4320-ae81-d7c48257e460: ' + [System.BitConverter]::ToString([System.BitConverter]::GetBytes($ARGB));
$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

function displayTooltip {
    try {
        # コンテキスト作成
        $appContext = New-Object System.Windows.Forms.ApplicationContext;

        ####################################################
        # 通知アイコン作成
        ####################################################
        $pwshPath = Get-Process -id $pid | Select-Object -ExpandProperty Path
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($pwshPath)    
        $notifyIcon = [System.Windows.Forms.NotifyIcon]@{
            Icon           = $icon;
            Text           = "CardWirthScenario DownloadWatcher";
            BalloonTipIcon = 'None';
        };

        ####################################################
        # アイコン左クリック時のイベントを設定
        ####################################################
        $notifyIcon.add_Click( {
                if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    try {
                        $notifyIcon.BalloonTipText = (Get-Date);
                        $notifyIcon.ShowBalloonTip(5000);
                    }
                    catch {
                        $notifyIcon.BalloonTipText = $_.ToString();
                        $notifyIcon.ShowBalloonTip(5000);
                    }
                }
            });

        ####################################################
        # アイコン右クリック時のコンテキストメニューの設定
        ####################################################
        $menuItem_exit = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Exit' };
        $menuItem_openFolder1 = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'ダウンロードフォルダを開く' };
        $menuItem_openFolder2 = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'シナリオフォルダを開く' };
        $menuItem_openFolder3 = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Nextシナリオフォルダを開く' };
        
        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_openFolder1);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_openFolder2);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_openFolder3);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
                $appContext.ExitThread();
            });
        $menuItem_openFolder1.add_Click( {
                Invoke-Item -Path ([WildcardPattern]::Escape($WATCH_DIRECTORY_PATH))
            });
        $menuItem_openFolder2.add_Click( {
                Invoke-Item -Path ([WildcardPattern]::Escape($CARDWIRTH_SCENARIO_DIRECTORY_PATH))
            });
        $menuItem_openFolder3.add_Click( {
                Invoke-Item -Path ([WildcardPattern]::Escape($CARDWIRTHNEXT_SCENARIO_DIRECTORY_PATH))
            });

        ####################################################
        # フォルダ監視の設定
        ####################################################

        # タイマーイベントがないと何故かRegister-ObjectEventが動かないので仕方なく設定している
        $timer = New-Object Windows.Forms.Timer
        $timer.Enabled = $true
        $timer.Add_Tick( {
                $timer.Stop()
                # 出力を捨てる
                Write-Output "タイマーイベント" > $null
                $timer.Interval = $TIMER_INTERVAL
                $timer.Start()
            })
        $timer.Interval = 1
        $timer.Start()

        # FileSystemWatcherの作成
        $wait = New-Object System.IO.FileSystemWatcher
        $wait.NotifyFilter = [IO.NotifyFilters]::LastWrite
        $wait.Path = $WATCH_DIRECTORY_PATH
        $wait.Filter = "*.*"

        # Register-ObjectEventの作成
        Register-ObjectEvent -InputObject $wait -SourceIdentifier "cw_scenario_download_watcher" -EventName "Changed" -Action {
            $t = $Event.TimeGenerated
            $f = $Event.SourceEventArgs
            [string]$lated_time = $t.ToString("yy/MM/dd HH:mm:ss.f")
            if ($lated_time -ne $chk_time) {
                
                Write-Host ("FilePath:" + $f.FullPath)
                $extension = [System.IO.Path]::GetExtension($f.FullPath).ToLower();
                Write-Host ("extension:" + $extension)
                if ($extension -in @(".cab", ".wsn", ".zip")) {
                    write-host "シナリオかもしれない"
                    if (Test-CardWirthScenario -LiteralPath $f.FullPath) {
                        $Scenario = Get-CardWirthScenario -LiteralPath $f.FullPath
                        write-host "シナリオ名:$($Scenario.Name)"
                        if ($Scenario.ScenarioType.ToString() -ne "Next") {
                            Write-Host ("シナリオフォルダに移動します")
                            Move-Item -LiteralPath $f.FullPath -Destination $Event.MessageData.CardWirthScenarioFolder
                            Write-Host ("シナリオフォルダに移動しました")
                            $Event.MessageData.notifyIcon.BalloonTipTitle = 'CardWirthシナリオのダウンロード';
                            $Event.MessageData.notifyIcon.BalloonTipText = "シナリオフォルダに移動しました。`nシナリオ名:$($Scenario.Name)`n作者:$($Scenario.Author)";
                            $Event.MessageData.notifyIcon.ShowBalloonTip(5000);
                        }
                        else {
                            Write-Host ("Nextシナリオフォルダに移動します")
                            Move-Item -LiteralPath $f.FullPath -Destination $Event.MessageData.CardWirthNextScenarioFolder
                            Write-Host ("Nextシナリオフォルダに移動しました")
                            $Event.MessageData.notifyIcon.BalloonTipTitle = 'CardWirthNextシナリオのダウンロード';
                            $Event.MessageData.notifyIcon.BalloonTipText = "Nextシナリオフォルダに移動しました。`nシナリオ名:$($Scenario.Name)`n作者:$($Scenario.Author)";
                            $Event.MessageData.notifyIcon.ShowBalloonTip(5000);
                        }
                    }
                }
                [string]$chk_time = $lated_time
            }
        } -MessageData @{"CardWirthScenarioFolder" = $CARDWIRTH_SCENARIO_DIRECTORY_PATH; "CardWirthNextScenarioFolder" = $CARDWIRTHNEXT_SCENARIO_DIRECTORY_PATH; "notifyIcon" = $notifyIcon }

        # 表示
        $notifyIcon.Visible = $true;
        [void][System.Windows.Forms.Application]::Run($appContext);
        $notifyIcon.Visible = $false;
    
    }
    finally {
        $notifyIcon.Dispose();
        $mutex.ReleaseMutex();
    }
}

# タスクバー非表示
function hiddenTaskber {
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
}
try {
    # タイトルバーの書き換え
    $Host.UI.RawUI.WindowTitle = "CardWirthScenario DownloadWatcher"
    # 多重起動チェック
    if ($mutex.WaitOne(0, $false)) {
        Write-Host "hiddenTaskber"
        hiddenTaskber
        Write-Host "displayTooltip"
        displayTooltip
        $retcode = 0;
    }
    else {
        $retcode = 255;
    }
}
finally {
    $mutex.Dispose();
}
exit $retcode;