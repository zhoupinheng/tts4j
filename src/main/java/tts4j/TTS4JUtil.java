/*
 * Copyright (c) 2018-2028 Github tts4j Project.
 * All rights reserved. Originator: Bruce.Zhou.
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
 */

package tts4j;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TTS4JUtil {
  final static Logger log = LoggerFactory.getLogger(TTS4JUtil.class);

  public static void speak(int volume, int rate, String words) {

    prepareNativeLibrary();

    speakText(volume, rate, words);

  }

  /**
   * @param volume
   *          0-100
   * @param rate
   *          -10 - +10
   * @param words
   *          words to speak
   */
  private static void speakText(int volume, int rate, String words) {
    ActiveXComponent sap = new ActiveXComponent("Sapi.SpVoice");
    Dispatch sapo = (Dispatch) sap.getObject();
    try {

      sap.setProperty("Volume", new Variant(volume));
      sap.setProperty("Rate", new Variant(rate));

      WINTYPE winType = getWinddowsType();
      if (winType != WINTYPE.WIN7) {
        Variant allVoices = Dispatch.call(sapo, "GetVoices");
        Dispatch dispVoices = allVoices.toDispatch();
        Dispatch setvoice = Dispatch.call(dispVoices, "Item", new Variant(1)).toDispatch();
        ActiveXComponent setvoiceActivex = new ActiveXComponent(setvoice);

        Dispatch.call(setvoiceActivex, "GetDescription");
      }

      Dispatch.call(sapo, "Speak", new Variant(words));

    } catch (Exception e) {
      log.error("speak error", e);
    } finally {
      sapo.safeRelease();
      sap.safeRelease();
    }
  }

  /**
   * copy jacob.dll to java.library.path
   */
  private static void prepareNativeLibrary() {
    try {
      String libraryStr = System.getProperty("java.library.path");
      String[] paths = libraryStr.split(File.pathSeparator);
      String filePath = null;
      File jacobFile = null;
      for (String dir : paths) {

        if (dir != null && dir.trim().length() > 0) {
          File libDir = new File(dir);
          if (libDir.exists() && libDir.isDirectory()) {
            if ("64".equals(System.getProperty("sun.arch.data.model"))) {
              filePath = libDir.getAbsolutePath() + File.separator + "jacob-1.19-x64.dll";
            } else {

              filePath = libDir.getAbsolutePath() + File.separator + "jacob-1.19-x86.dll";
            }

            jacobFile = new File(filePath);
            if (jacobFile.exists() && jacobFile.canRead()) {
              // dll exist in library path
              if (log.isInfoEnabled()) {
                log.info("jacob dll file exist,the path is :" + jacobFile.getAbsolutePath());
              }
              break;
            } else {
              if (libDir.canWrite()) {

                InputStream in = null;
                if ("64".equals(System.getProperty("sun.arch.data.model"))) {
                  in = TTS4JUtil.class.getResourceAsStream("/jacob-1.19-x64.dll");
                } else {
                  in = TTS4JUtil.class.getResourceAsStream("/jacob-1.19-x86.dll");
                }

                FileOutputStream out = new FileOutputStream(jacobFile);

                int i;
                byte[] buf = new byte[1024];
                try {
                  while ((i = in.read(buf)) != -1) {
                    out.write(buf, 0, i);
                  }
                  if (log.isInfoEnabled()) {
                    log.info("generate new jacob file, the path is :" + jacobFile.getAbsolutePath());
                  }
                  break;
                } finally {
                  in.close();
                  out.close();
                }

              }
            }

          }
        }
      }

      if (jacobFile != null && jacobFile.exists() && jacobFile.canRead()) {
        System.load(jacobFile.getAbsolutePath());
      }

    } catch (Exception e) {
      log.error("load jni error!", e);
      e.printStackTrace();
    }
  }

  private static WINTYPE getWinddowsType() {
    String osName = System.getProperty("os.name");
    String version = System.getProperty("os.version");

    if ("Windows 7".equalsIgnoreCase(osName.trim())) {
      return WINTYPE.WIN7;
    } else if (osName.startsWith("Windows 8.1")) {
      return WINTYPE.WIN10;
    } else {
      return WINTYPE.WIN8;
    }
  }

  enum WINTYPE {
    WIN7("WIN7"), WIN8("WIN8"), WIN10("WIN10");

    private WINTYPE(String desc) {
      this.description = desc;
    }

    public String toString() {
      return description;
    }

    private String description;
  }
  
}
