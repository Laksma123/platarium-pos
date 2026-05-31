package com.platarium.pos;

import android.Manifest;
import android.content.pm.PackageManager;
import android.media.MediaPlayer;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.View;
import android.view.WindowManager;
import android.view.animation.AlphaAnimation;
import android.view.animation.Animation;
import android.webkit.PermissionRequest;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;
import android.widget.VideoView;
import android.content.Intent;
import android.webkit.ValueCallback;
import android.content.ActivityNotFoundException;
import android.util.Base64;
import java.io.File;
import java.io.FileOutputStream;
import android.os.Environment;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.view.WindowCompat;
import androidx.core.view.WindowInsetsCompat;
import androidx.core.view.WindowInsetsControllerCompat;

public class MainActivity extends AppCompatActivity {

    // ==========================================
    // CONFIGURATION - GANTI DENGAN URL WEBAPP ANDA
    // ==========================================
    private static final String TARGET_URL = "https://laksma123.github.io/platarium-pos/";
    
    private static final int REQUEST_CAMERA_PERMISSION = 101;
    private static final int FILECHOOSER_RESULTCODE = 102;
    private ValueCallback<Uri[]> mFilePathCallback;
    
    private WebView webView;
    private VideoView splashVideoView;
    private boolean isTransitioned = false;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // 1. Inisialisasi Immersive Fullscreen Mode
        setImmersiveFullscreen();

        // 2. Inisialisasi WebView
        webView = findViewById(R.id.webView);
        setupWebView();

        // 3. Inisialisasi VideoView & Putar Video Boot
        splashVideoView = findViewById(R.id.splashVideoView);
        setupAndPlaySplashVideo();

        // 4. Periksa dan minta izin kamera runtime (Android 6.0+) serta mulai load WebView di background
        checkAndRequestCameraPermission();
    }

    @Override
    protected void onResume() {
        super.onResume();
        // Pastikan mode fullscreen tetap aktif saat aplikasi kembali fokus
        setImmersiveFullscreen();
    }

    /**
     * Mengatur Tampilan Fullscreen & Immersive Mode Profesional
     */
    private void setImmersiveFullscreen() {
        WindowCompat.setDecorFitsSystemWindows(getWindow(), false);
        WindowInsetsControllerCompat controller = new WindowInsetsControllerCompat(getWindow(), getWindow().getDecorView());
        controller.hide(WindowInsetsCompat.Type.statusBars() | WindowInsetsCompat.Type.navigationBars());
        controller.setSystemBarsBehavior(WindowInsetsControllerCompat.BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE);

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.P) {
            WindowManager.LayoutParams layoutParams = getWindow().getAttributes();
            layoutParams.layoutInDisplayCutoutMode = WindowManager.LayoutParams.LAYOUT_IN_DISPLAY_CUTOUT_MODE_SHORT_EDGES;
            getWindow().setAttributes(layoutParams);
        }
    }

    /**
     * Memutar Video Splash Boot di Foreground & Menangani Transisi
     */
    private void setupAndPlaySplashVideo() {
        try {
            // Arahkan ke file res/raw/splash_video.mp4
            String videoPath = "android.resource://" + getPackageName() + "/" + R.raw.splash_video;
            Uri videoUri = Uri.parse(videoPath);
            splashVideoView.setVideoURI(videoUri);

            // Terapkan Center Crop Dinamis ketika video siap diputar (OnPrepared)
            splashVideoView.setOnPreparedListener(new MediaPlayer.OnPreparedListener() {
                @Override
                public void onPrepared(MediaPlayer mp) {
                    // Dapatkan ukuran asli video
                    int videoWidth = mp.getVideoWidth();
                    int videoHeight = mp.getVideoHeight();

                    // Dapatkan ukuran kontainer layar VideoView
                    int viewWidth = splashVideoView.getWidth();
                    int viewHeight = splashVideoView.getHeight();

                    if (videoWidth > 0 && videoHeight > 0 && viewWidth > 0 && viewHeight > 0) {
                        float videoRatio = (float) videoWidth / (float) videoHeight;
                        float viewRatio = (float) viewWidth / (float) viewHeight;

                        android.view.ViewGroup.LayoutParams lp = splashVideoView.getLayoutParams();

                        if (viewRatio > videoRatio) {
                            // Layar lebih lebar dibanding rasio video -> potong atas bawah (lebar fit, tinggi melar keluar)
                            lp.width = viewWidth;
                            lp.height = (int) (viewWidth / videoRatio);
                        } else {
                            // Layar lebih tinggi dibanding rasio video -> potong kiri kanan (tinggi fit, lebar melar keluar)
                            lp.height = viewHeight;
                            lp.width = (int) (viewHeight * videoRatio);
                        }
                        splashVideoView.setLayoutParams(lp);
                    }
                    
                    // Mulai putar video setelah layout disesuaikan
                    splashVideoView.start();
                }
            });

            // Ketika video selesai diputar, lakukan transisi fade out
            splashVideoView.setOnCompletionListener(new MediaPlayer.OnCompletionListener() {
                @Override
                public void onCompletion(MediaPlayer mp) {
                    fadeOutSplash();
                }
            });

            // Penanganan jika video error (misalnya file rusak, format tidak didukung, atau masih menggunakan file placeholder)
            splashVideoView.setOnErrorListener(new MediaPlayer.OnErrorListener() {
                @Override
                public boolean onError(MediaPlayer mp, int what, int extra) {
                    // Langsung hilangkan splash screen agar aplikasi tidak stuck hitam
                    fadeOutSplash();
                    return true;
                }
            });

            // Safety Handler: Batasan waktu maksimum 12 detik sebagai pengaman
            new Handler(Looper.getMainLooper()).postDelayed(new Runnable() {
                @Override
                public void run() {
                    fadeOutSplash();
                }
            }, 12000);

        } catch (Exception e) {
            e.printStackTrace();
            fadeOutSplash();
        }
    }

    /**
     * Efek Memudarkan (Fade Out) Video Splash Screen Secara Lembut (500ms)
     */
    private synchronized void fadeOutSplash() {
        if (isTransitioned) return; // Mencegah pemanggilan ganda
        isTransitioned = true;

        if (splashVideoView.getVisibility() == View.VISIBLE) {
            // Hentikan video untuk menghemat resource
            try {
                splashVideoView.stopPlayback();
            } catch (Exception ignored) {}

            // Buat animasi memudar (Alpha dari 1.0 menjadi 0.0)
            AlphaAnimation fadeOut = new AlphaAnimation(1.0f, 0.0f);
            fadeOut.setDuration(500); // 500 milidetik
            fadeOut.setAnimationListener(new Animation.AnimationListener() {
                @Override
                public void onAnimationStart(Animation animation) {}

                @Override
                public void onAnimationEnd(Animation animation) {
                    splashVideoView.setVisibility(View.GONE);
                    
                    // FAKTOR UTAMA: Hapus VideoView dari parent secara permanen agar surface hardware-nya dihancurkan
                    android.view.ViewGroup parent = (android.view.ViewGroup) splashVideoView.getParent();
                    if (parent != null) {
                        parent.removeView(splashVideoView);
                    }

                    // Paksa WebView untuk muncul paling depan, aktifkan, dan gambar ulang dirinya
                    webView.bringToFront();
                    webView.setVisibility(View.VISIBLE);
                    webView.onResume(); // Pastikan siklus hidup WebView berjalan kembali
                    webView.invalidate(); // Paksa penggambaran ulang (redraw)
                }

                @Override
                public void onAnimationRepeat(Animation animation) {}
            });

            splashVideoView.startAnimation(fadeOut);
        }
    }

    /**
     * Konfigurasi Parameter WebView Secara Profesional (Latar Belakang Dimuat Senyap)
     */
    private void setupWebView() {
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setCacheMode(WebSettings.LOAD_DEFAULT);
        settings.setMediaPlaybackRequiresUserGesture(false);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);

        webView.setBackgroundColor(0); // Buat background WebView transparan agar tidak ada warna putih bocor saat resize/keyboard muncul
        String defaultUserAgent = settings.getUserAgentString();
        settings.setUserAgentString(defaultUserAgent + " PlatariumPOSAndroidWrapper/1.0");

        webView.addJavascriptInterface(new WebAppInterface(this), "AndroidDownloader");

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                view.loadUrl(url);
                return true;
            }
        });

        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onPermissionRequest(final PermissionRequest request) {
                MainActivity.this.runOnUiThread(new Runnable() {
                    @Override
                    public void run() {
                        String[] resources = request.getResources();
                        for (String resource : resources) {
                            if (PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(resource) ||
                                PermissionRequest.RESOURCE_AUDIO_CAPTURE.equals(resource)) {
                                request.grant(resources);
                                return;
                            }
                        }
                        request.deny();
                    }
                });
            }

            @Override
            public boolean onShowFileChooser(WebView webView, ValueCallback<Uri[]> filePathCallback, WebChromeClient.FileChooserParams fileChooserParams) {
                if (mFilePathCallback != null) {
                    mFilePathCallback.onReceiveValue(null);
                }
                mFilePathCallback = filePathCallback;

                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                intent.setType("image/*");

                try {
                    startActivityForResult(Intent.createChooser(intent, "Pilih Gambar"), FILECHOOSER_RESULTCODE);
                } catch (ActivityNotFoundException e) {
                    mFilePathCallback = null;
                    Toast.makeText(MainActivity.this, "Tidak dapat membuka file manager", Toast.LENGTH_LONG).show();
                    return false;
                }
                return true;
            }
        });
    }

    /**
     * Runtime Permission Android 6.0+
     */
    private void checkAndRequestCameraPermission() {
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.CAMERA)
                != PackageManager.PERMISSION_GRANTED) {
            
            ActivityCompat.requestPermissions(this,
                    new String[]{Manifest.permission.CAMERA},
                    REQUEST_CAMERA_PERMISSION);
        } else {
            webView.loadUrl(TARGET_URL);
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == REQUEST_CAMERA_PERMISSION) {
            if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                webView.loadUrl(TARGET_URL);
            } else {
                Toast.makeText(this, "Izin kamera ditolak. Fitur scanner kehadiran tidak dapat berfungsi.", Toast.LENGTH_LONG).show();
                webView.loadUrl(TARGET_URL);
            }
        }
    }

    @Override
    public void onBackPressed() {
        if (webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == FILECHOOSER_RESULTCODE) {
            if (mFilePathCallback == null) return;
            Uri[] results = null;
            if (resultCode == RESULT_OK && data != null) {
                String dataString = data.getDataString();
                if (dataString != null) {
                    results = new Uri[]{Uri.parse(dataString)};
                } else if (data.getClipData() != null) {
                    int count = data.getClipData().getItemCount();
                    results = new Uri[count];
                    for (int i = 0; i < count; i++) {
                        results[i] = data.getClipData().getItemAt(i).getUri();
                    }
                }
            }
            mFilePathCallback.onReceiveValue(results);
            mFilePathCallback = null;
        } else {
            super.onActivityResult(requestCode, resultCode, data);
        }
    }

    public class WebAppInterface {
        android.content.Context mContext;
        WebAppInterface(android.content.Context c) {
            mContext = c;
        }

        @android.webkit.JavascriptInterface
        public void saveBase64Pdf(String base64, String filename) {
            try {
                byte[] pdfAsBytes = Base64.decode(base64, Base64.DEFAULT);
                File path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
                File file = new File(path, filename);
                FileOutputStream os = new FileOutputStream(file, false);
                os.write(pdfAsBytes);
                os.flush();
                os.close();
                
                new Handler(Looper.getMainLooper()).post(new Runnable() {
                    @Override
                    public void run() {
                        Toast.makeText(mContext, "SOP Disimpan di folder Download!", Toast.LENGTH_LONG).show();
                    }
                });
            } catch (Exception e) {
                new Handler(Looper.getMainLooper()).post(new Runnable() {
                    @Override
                    public void run() {
                        Toast.makeText(mContext, "Gagal simpan SOP: " + e.getMessage(), Toast.LENGTH_LONG).show();
                    }
                });
            }
        }
    }
}
