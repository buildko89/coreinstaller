using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Threading.Tasks;

namespace CustomAction
{
  [RunInstaller(true)]
  public partial class Installer1 : System.Configuration.Install.Installer
  {
    // インストール時の動作関数
    public override void Install(System.Collections.IDictionary stateSaver)
    {
      // Install後の動作
      base.Install(stateSaver);

      // 環境変数PATHの追加
      string currentPath;
      currentPath = System.Environment.GetEnvironmentVariable("path", System.EnvironmentVariableTarget.User);
      string installPath = this.Context.Parameters["InstallPath"];
      string path = installPath + @"\bin;";

#if DEBUG
      System.Windows.Forms.MessageBox.Show(installPath);
#endif

      if (currentPath == null)
      {
        currentPath = path;
      }
      else if (currentPath.EndsWith(";"))
      {
        currentPath += path;
      }
      else
      {
        currentPath += ";" + path;
      }

      // 環境変数PATHを設定する
      System.Environment.SetEnvironmentVariable("path", currentPath, System.EnvironmentVariableTarget.User);

      //hakoniwa core config pathの設定
      string configpath = installPath + @"\config\cpp_core_config.json";
      System.Environment.SetEnvironmentVariable("HAKO_CONFIG_PATH", configpath, System.EnvironmentVariableTarget.User);

      //hakoniwa core Library pathの設定
      string libpath = installPath + @"\lib";
      System.Environment.SetEnvironmentVariable("HAKOCORE_LIB_PATH", libpath, System.EnvironmentVariableTarget.User);

      //hakoniwa core Python pathの設定
      string pythonPath;

      pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
      string hakopypath = installPath + @"\lib\py;";
#if DEBUG
      System.Windows.Forms.MessageBox.Show(hakopypath);
#endif

      if (pythonPath == null)
      {
        pythonPath = hakopypath;
      }
      else if (pythonPath.EndsWith(";"))
      {
        pythonPath += hakopypath;
      }
      else
      {
        pythonPath += ";" + hakopypath;
      }

      // 環境変数PYTHONPATHを設定する
      System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
      System.Windows.Forms.MessageBox.Show("Install");
#endif

      // hakoniwa-pduのpip installを実行
      try
      {
        var psi = new System.Diagnostics.ProcessStartInfo();
        psi.FileName = "powershell.exe";
        psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"python -m pip install hakoniwa-pdu\"";
        psi.UseShellExecute = false;
        psi.RedirectStandardOutput = true;
        psi.RedirectStandardError = true;
        psi.CreateNoWindow = true;
        psi.EnvironmentVariables["PATH"] = System.Environment.GetEnvironmentVariable("PATH", System.EnvironmentVariableTarget.User);

        using (var process = System.Diagnostics.Process.Start(psi))
        {
          string output = process.StandardOutput.ReadToEnd();
          string error = process.StandardError.ReadToEnd();
          process.WaitForExit();

          if (!string.IsNullOrEmpty(error))
          {
            System.Windows.Forms.MessageBox.Show("pip install エラー:\n" + error, "エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
          }
          else
          {
            System.Windows.Forms.MessageBox.Show("pip install 成功:\n" + output, "成功", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
          }
        }
      }
      catch (Exception ex)
      {
        System.Windows.Forms.MessageBox.Show("PowerShell 実行例外:\n" + ex.Message, "例外", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
      }
    }

    // インストールの状態を変更する動作関数
    public override void Commit(System.Collections.IDictionary savedState)
    {
      //変更時の動作
      base.Commit(savedState);
#if DEBUG
      System.Windows.Forms.MessageBox.Show("Commit");
#endif
    }

    // インストール失敗時の修復動作関数
    public override void Rollback(System.Collections.IDictionary savedState)
    {
      //修復動作
      base.Rollback(savedState);
#if DEBUG
      System.Windows.Forms.MessageBox.Show("Rollback");
#endif
    }

    // アンインストール時の動作関数
    public override void Uninstall(System.Collections.IDictionary savedState)
    {
      //Un-install動作
      base.Uninstall(savedState);

      // 環境変数PATHを編集
      string currentPath;
      currentPath = System.Environment.GetEnvironmentVariable("path", System.EnvironmentVariableTarget.User);
      string installPath = this.Context.Parameters["InstallPath"];
      string path = installPath + @"\bin;";
      currentPath = currentPath.Replace(path, "");

      // 環境変数PATHから削除する
      System.Environment.SetEnvironmentVariable("path", currentPath, System.EnvironmentVariableTarget.User);

      // hakoniwa core関連のPATHを削除する
      System.Environment.SetEnvironmentVariable("HAKOCORE_LIB_PATH", "", System.EnvironmentVariableTarget.User);
      System.Environment.SetEnvironmentVariable("HAKO_CONFIG_PATH", "", System.EnvironmentVariableTarget.User);

      // PYTHONPATHから箱庭コア関連を消す
      string pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
      string hakopypath = installPath + @"\lib\py;";
      pythonPath = pythonPath.Replace(hakopypath, "");

#if DEBUG
      System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
      // PYTHONPATHから箱庭コア関連を削除
      System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
      System.Windows.Forms.MessageBox.Show("Uninstall");
#endif
    }
  }
}
