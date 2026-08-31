package com.kcc.hdi.lotgenerator;

import android.app.Activity;
import android.content.ClipData;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Color;
import android.graphics.Matrix;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.view.View;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import androidx.core.content.FileProvider;
import androidx.exifinterface.media.ExifInterface;

import com.google.mlkit.vision.common.InputImage;
import com.google.mlkit.vision.text.TextRecognition;
import com.google.mlkit.vision.text.TextRecognizer;
import com.google.mlkit.vision.text.latin.TextRecognizerOptions;
import com.google.zxing.integration.android.IntentIntegrator;
import com.google.zxing.integration.android.IntentResult;

import org.json.JSONObject;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class MainActivity extends Activity {
    private static final String APP_URL = "file:///android_asset/www/index.html";
    private static final int PANEL_PHOTO_REQUEST = 7001;
    private static final String FILE_PROVIDER_SUFFIX = ".fileprovider";

    private WebView webView;
    private String activeScanTarget = "lot";
    private String pendingExpectedValue = "";
    private File pendingPanelPhoto;
    private Uri pendingPanelPhotoUri;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        getWindow().setStatusBarColor(Color.rgb(18, 60, 105));
        getWindow().setNavigationBarColor(Color.rgb(18, 60, 105));

        webView = new WebView(this);
        webView.setBackgroundColor(Color.rgb(243, 246, 249));
        webView.setOverScrollMode(View.OVER_SCROLL_NEVER);
        setContentView(webView);

        configureWebView(webView);
        webView.addJavascriptInterface(new AndroidScannerBridge(), "AndroidScanner");

        if (savedInstanceState == null) {
            webView.loadUrl(APP_URL);
        } else {
            webView.restoreState(savedInstanceState);
        }
    }

    private void startBarcodeScan(String target) {
        activeScanTarget = target;
        IntentIntegrator integrator = new IntentIntegrator(this);
        if ("management".equals(target)) {
            integrator.setPrompt("관리번호 바코드를 화면 안에 맞춰주세요");
        } else {
            integrator.setPrompt("LOT NO 바코드를 화면 안에 맞춰주세요");
        }
        integrator.setBeepEnabled(true);
        integrator.setBarcodeImageEnabled(false);
        integrator.setOrientationLocked(false);
        integrator.setCameraId(0);
        integrator.setDesiredBarcodeFormats(IntentIntegrator.ALL_CODE_TYPES);
        integrator.initiateScan();
    }

    private void startPanelPhotoVerification(String expectedValue) {
        String normalizedExpected = expectedValue == null
                ? ""
                : expectedValue.trim().toUpperCase(Locale.ROOT);
        if (!normalizedExpected.matches("^[A-Z0-9]{13}$")) {
            notifyPanelOcrFailed("생성된 각인값이 올바르지 않습니다. 입력값을 다시 확인해 주세요.");
            return;
        }

        deletePendingPanelPhoto();
        pendingExpectedValue = normalizedExpected;

        try {
            File photoDirectory = new File(getCacheDir(), "panel_photos");
            if (!photoDirectory.exists() && !photoDirectory.mkdirs()) {
                notifyPanelOcrFailed("사진 임시 폴더를 만들지 못했습니다.");
                return;
            }

            pendingPanelPhoto = File.createTempFile("first_panel_", ".jpg", photoDirectory);
            pendingPanelPhotoUri = FileProvider.getUriForFile(
                    this,
                    getPackageName() + FILE_PROVIDER_SUFFIX,
                    pendingPanelPhoto
            );

            Intent cameraIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
            cameraIntent.putExtra(MediaStore.EXTRA_OUTPUT, pendingPanelPhotoUri);
            cameraIntent.setClipData(ClipData.newRawUri("초도품 각인 사진", pendingPanelPhotoUri));
            cameraIntent.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_READ_URI_PERMISSION);

            if (cameraIntent.resolveActivity(getPackageManager()) == null) {
                deletePendingPanelPhoto();
                notifyPanelOcrFailed("사진을 촬영할 카메라 앱을 찾지 못했습니다.");
                return;
            }
            startActivityForResult(cameraIntent, PANEL_PHOTO_REQUEST);
        } catch (IOException | IllegalArgumentException error) {
            deletePendingPanelPhoto();
            notifyPanelOcrFailed("사진 촬영 준비 중 오류가 발생했습니다: " + safeMessage(error));
        }
    }

    private void processPanelPhoto() {
        if (pendingPanelPhotoUri == null || pendingPanelPhoto == null || !pendingPanelPhoto.exists()) {
            notifyPanelOcrFailed("촬영한 사진을 찾지 못했습니다. 다시 촬영해 주세요.");
            deletePendingPanelPhoto();
            return;
        }

        final InputImage image;
        try {
            image = InputImage.fromFilePath(this, pendingPanelPhotoUri);
        } catch (IOException error) {
            notifyPanelOcrFailed("촬영한 사진을 읽지 못했습니다: " + safeMessage(error));
            deletePendingPanelPhoto();
            return;
        }

        TextRecognizer recognizer = TextRecognition.getClient(TextRecognizerOptions.DEFAULT_OPTIONS);
        recognizer.process(image)
                .addOnSuccessListener(result -> {
                    String rawText = result.getText() == null ? "" : result.getText().trim();
                    String normalizedText = normalizeRecognizedText(rawText);
                    if (isExpectedText(normalizedText)) {
                        finishPanelRecognition(recognizer, null, rawText, normalizedText, true);
                    } else {
                        startPreprocessedRecognition(recognizer, rawText, normalizedText);
                    }
                })
                .addOnFailureListener(error -> startPreprocessedRecognition(recognizer, "", ""));
    }

    private void startPreprocessedRecognition(
            TextRecognizer recognizer,
            String firstRawText,
            String firstNormalizedText
    ) {
        final List<Bitmap> variants;
        try {
            variants = createOcrVariants(pendingPanelPhoto);
        } catch (IOException | RuntimeException error) {
            finishPanelRecognition(
                    recognizer,
                    null,
                    firstRawText,
                    firstNormalizedText,
                    false
            );
            return;
        }
        recognizeVariant(
                recognizer,
                variants,
                0,
                firstRawText,
                firstNormalizedText
        );
    }

    private void recognizeVariant(
            TextRecognizer recognizer,
            List<Bitmap> variants,
            int index,
            String bestRawText,
            String bestNormalizedText
    ) {
        if (index >= variants.size()) {
            finishPanelRecognition(
                    recognizer,
                    variants,
                    bestRawText,
                    bestNormalizedText,
                    false
            );
            return;
        }

        Bitmap bitmap = variants.get(index);
        recognizer.process(InputImage.fromBitmap(bitmap, 0))
                .addOnSuccessListener(result -> {
                    String rawText = result.getText() == null ? "" : result.getText().trim();
                    String normalizedText = normalizeRecognizedText(rawText);
                    if (isExpectedText(normalizedText)) {
                        finishPanelRecognition(
                                recognizer,
                                variants,
                                rawText,
                                normalizedText,
                                true
                        );
                        return;
                    }

                    String nextRawText = bestRawText;
                    String nextNormalizedText = bestNormalizedText;
                    if (normalizedText.length() > bestNormalizedText.length()) {
                        nextRawText = rawText;
                        nextNormalizedText = normalizedText;
                    }
                    recognizeVariant(
                            recognizer,
                            variants,
                            index + 1,
                            nextRawText,
                            nextNormalizedText
                    );
                })
                .addOnFailureListener(error -> recognizeVariant(
                        recognizer,
                        variants,
                        index + 1,
                        bestRawText,
                        bestNormalizedText
                ));
    }

    private void finishPanelRecognition(
            TextRecognizer recognizer,
            List<Bitmap> variants,
            String rawText,
            String normalizedText,
            boolean matched
    ) {
        if (variants != null) {
            for (Bitmap bitmap : variants) {
                if (bitmap != null && !bitmap.isRecycled()) bitmap.recycle();
            }
        }
        recognizer.close();
        notifyPanelOcrCompleted(rawText, normalizedText, matched);
        deletePendingPanelPhoto();
        pendingExpectedValue = "";
    }

    private boolean isExpectedText(String normalizedText) {
        return !pendingExpectedValue.isEmpty()
                && normalizedText != null
                && normalizedText.contains(pendingExpectedValue);
    }

    private List<Bitmap> createOcrVariants(File photoFile) throws IOException {
        BitmapFactory.Options bounds = new BitmapFactory.Options();
        bounds.inJustDecodeBounds = true;
        BitmapFactory.decodeFile(photoFile.getAbsolutePath(), bounds);
        if (bounds.outWidth <= 0 || bounds.outHeight <= 0) {
            throw new IOException("사진 크기를 확인하지 못했습니다.");
        }

        BitmapFactory.Options options = new BitmapFactory.Options();
        options.inSampleSize = 1;
        while (bounds.outWidth / options.inSampleSize > 2600
                || bounds.outHeight / options.inSampleSize > 2600) {
            options.inSampleSize *= 2;
        }
        options.inPreferredConfig = Bitmap.Config.ARGB_8888;

        Bitmap decoded = BitmapFactory.decodeFile(photoFile.getAbsolutePath(), options);
        if (decoded == null) throw new IOException("사진 Bitmap을 만들지 못했습니다.");

        Bitmap oriented = applyExifOrientation(decoded, photoFile);
        if (oriented != decoded && !decoded.isRecycled()) decoded.recycle();

        int sourceWidth = oriented.getWidth();
        int targetWidth = Math.max(1600, Math.min(2200, sourceWidth));
        Bitmap base = oriented;
        if (sourceWidth != targetWidth) {
            int targetHeight = Math.max(1, Math.round(
                    oriented.getHeight() * (targetWidth / (float) sourceWidth)
            ));
            base = Bitmap.createScaledBitmap(oriented, targetWidth, targetHeight, true);
            if (base != oriented && !oriented.isRecycled()) oriented.recycle();
        }

        List<Bitmap> variants = new ArrayList<>();
        variants.add(base);
        variants.add(toHighContrast(base));
        variants.add(toBinary(base, 90));
        variants.add(toBinary(base, 120));
        variants.add(toBinary(base, 150));
        variants.add(toBinary(base, 180));
        variants.add(toBinary(base, 210));
        return variants;
    }

    private Bitmap applyExifOrientation(Bitmap source, File photoFile) throws IOException {
        ExifInterface exif = new ExifInterface(photoFile);
        int orientation = exif.getAttributeInt(
                ExifInterface.TAG_ORIENTATION,
                ExifInterface.ORIENTATION_NORMAL
        );
        Matrix matrix = new Matrix();
        switch (orientation) {
            case ExifInterface.ORIENTATION_FLIP_HORIZONTAL:
                matrix.setScale(-1, 1);
                break;
            case ExifInterface.ORIENTATION_ROTATE_180:
                matrix.setRotate(180);
                break;
            case ExifInterface.ORIENTATION_FLIP_VERTICAL:
                matrix.setRotate(180);
                matrix.postScale(-1, 1);
                break;
            case ExifInterface.ORIENTATION_TRANSPOSE:
                matrix.setRotate(90);
                matrix.postScale(-1, 1);
                break;
            case ExifInterface.ORIENTATION_ROTATE_90:
                matrix.setRotate(90);
                break;
            case ExifInterface.ORIENTATION_TRANSVERSE:
                matrix.setRotate(-90);
                matrix.postScale(-1, 1);
                break;
            case ExifInterface.ORIENTATION_ROTATE_270:
                matrix.setRotate(-90);
                break;
            default:
                return source;
        }
        return Bitmap.createBitmap(
                source,
                0,
                0,
                source.getWidth(),
                source.getHeight(),
                matrix,
                true
        );
    }

    private Bitmap toHighContrast(Bitmap source) {
        int width = source.getWidth();
        int height = source.getHeight();
        int[] pixels = new int[width * height];
        source.getPixels(pixels, 0, width, 0, 0, width, height);
        for (int index = 0; index < pixels.length; index++) {
            int color = pixels[index];
            int gray = (int) (0.299 * ((color >> 16) & 0xff)
                    + 0.587 * ((color >> 8) & 0xff)
                    + 0.114 * (color & 0xff));
            int contrast = Math.max(0, Math.min(255, (gray - 128) * 2 + 128));
            pixels[index] = 0xff000000 | (contrast << 16) | (contrast << 8) | contrast;
        }
        Bitmap result = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888);
        result.setPixels(pixels, 0, width, 0, 0, width, height);
        return result;
    }

    private Bitmap toBinary(Bitmap source, int threshold) {
        int width = source.getWidth();
        int height = source.getHeight();
        int[] pixels = new int[width * height];
        source.getPixels(pixels, 0, width, 0, 0, width, height);
        for (int index = 0; index < pixels.length; index++) {
            int color = pixels[index];
            int gray = (int) (0.299 * ((color >> 16) & 0xff)
                    + 0.587 * ((color >> 8) & 0xff)
                    + 0.114 * (color & 0xff));
            int value = gray < threshold ? 0 : 255;
            pixels[index] = 0xff000000 | (value << 16) | (value << 8) | value;
        }
        Bitmap result = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888);
        result.setPixels(pixels, 0, width, 0, 0, width, height);
        return result;
    }

    private String normalizeRecognizedText(String value) {
        return value == null
                ? ""
                : value.toUpperCase(Locale.ROOT).replaceAll("[^A-Z0-9]", "");
    }

    private void notifyPanelOcrCompleted(String rawText, String normalizedText, boolean matched) {
        String script = "window.onPanelOcrCompleted("
                + JSONObject.quote(rawText) + ","
                + JSONObject.quote(normalizedText) + ","
                + matched + ");";
        evaluateJavascript(script);
    }

    private void notifyPanelOcrFailed(String message) {
        evaluateJavascript("window.onPanelOcrFailed(" + JSONObject.quote(message) + ");");
    }

    private void notifyPanelPhotoCancelled() {
        evaluateJavascript("window.onPanelPhotoCancelled();");
    }

    private void evaluateJavascript(String script) {
        if (webView != null) {
            webView.post(() -> {
                if (webView != null) webView.evaluateJavascript(script, null);
            });
        }
    }

    private String safeMessage(Exception error) {
        String message = error.getMessage();
        return message == null || message.trim().isEmpty() ? error.getClass().getSimpleName() : message;
    }

    private void deletePendingPanelPhoto() {
        if (pendingPanelPhoto != null && pendingPanelPhoto.exists()) {
            // 사진은 판정 직후 삭제하며 작업 이력에는 문자 결과만 저장합니다.
            pendingPanelPhoto.delete();
        }
        pendingPanelPhoto = null;
        pendingPanelPhotoUri = null;
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == PANEL_PHOTO_REQUEST) {
            if (resultCode == RESULT_OK) {
                processPanelPhoto();
            } else {
                deletePendingPanelPhoto();
                pendingExpectedValue = "";
                notifyPanelPhotoCancelled();
            }
            return;
        }

        IntentResult result = IntentIntegrator.parseActivityResult(requestCode, resultCode, data);
        if (result != null) {
            if (result.getContents() == null) {
                String safeTarget = JSONObject.quote(activeScanTarget);
                evaluateJavascript("window.onAndroidScanCancelled(" + safeTarget + ");");
            } else {
                String safeValue = JSONObject.quote(result.getContents());
                if ("management".equals(activeScanTarget)) {
                    evaluateJavascript("window.onAndroidManagementScanned(" + safeValue + ");");
                } else {
                    evaluateJavascript("window.onAndroidLotScanned(" + safeValue + ");");
                }
            }
            return;
        }
        super.onActivityResult(requestCode, resultCode, data);
    }

    private final class AndroidScannerBridge {
        @JavascriptInterface
        public void scanManagementBarcode() {
            runOnUiThread(() -> startBarcodeScan("management"));
        }

        @JavascriptInterface
        public void scanLotBarcode() {
            runOnUiThread(() -> startBarcodeScan("lot"));
        }

        @JavascriptInterface
        public void captureAndVerifyPanel(String expectedValue) {
            runOnUiThread(() -> startPanelPhotoVerification(expectedValue));
        }
    }

    private void configureWebView(WebView target) {
        WebSettings settings = target.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setAllowContentAccess(false);
        settings.setAllowFileAccess(true);
        settings.setDatabaseEnabled(false);
        settings.setGeolocationEnabled(false);
        settings.setMediaPlaybackRequiresUserGesture(true);
        settings.setSupportMultipleWindows(false);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);
        settings.setBlockNetworkLoads(true);

        target.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                return !url.startsWith("file:///android_asset/www/");
            }
        });
    }

    @Override
    protected void onSaveInstanceState(Bundle outState) {
        webView.saveState(outState);
        super.onSaveInstanceState(outState);
    }

    @Override
    protected void onDestroy() {
        deletePendingPanelPhoto();
        if (webView != null) {
            webView.loadUrl("about:blank");
            webView.stopLoading();
            webView.setWebViewClient(null);
            webView.destroy();
            webView = null;
        }
        super.onDestroy();
    }
}
