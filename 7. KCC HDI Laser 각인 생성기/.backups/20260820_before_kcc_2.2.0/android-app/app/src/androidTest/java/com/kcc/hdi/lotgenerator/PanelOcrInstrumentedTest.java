package com.kcc.hdi.lotgenerator;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import android.content.Context;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;

import androidx.test.ext.junit.runners.AndroidJUnit4;
import androidx.test.platform.app.InstrumentationRegistry;

import com.google.android.gms.tasks.Tasks;
import com.google.mlkit.vision.common.InputImage;
import com.google.mlkit.vision.text.Text;
import com.google.mlkit.vision.text.TextRecognition;
import com.google.mlkit.vision.text.TextRecognizer;
import com.google.mlkit.vision.text.latin.TextRecognizerOptions;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;

@RunWith(AndroidJUnit4.class)
public class PanelOcrInstrumentedTest {
    @Test
    public void recognizesProvidedFirstPanelPhoto() throws Exception {
        Context testContext = InstrumentationRegistry.getInstrumentation().getContext();
        Bitmap bitmap;
        try (InputStream stream = testContext.getAssets().open("panel_sample.png")) {
            bitmap = BitmapFactory.decodeStream(stream);
        }
        assertNotNull("테스트 사진을 읽지 못했습니다.", bitmap);

        List<Bitmap> variants = createVariants(bitmap);
        TextRecognizer recognizer = TextRecognition.getClient(TextRecognizerOptions.DEFAULT_OPTIONS);
        try {
            boolean matched = false;
            StringBuilder allResults = new StringBuilder();
            for (int index = 0; index < variants.size(); index++) {
                Text result = Tasks.await(
                        recognizer.process(InputImage.fromBitmap(variants.get(index), 0)),
                        30,
                        TimeUnit.SECONDS
                );
                String normalized = result.getText()
                        .toUpperCase(Locale.ROOT)
                        .replaceAll("[^A-Z0-9]", "");
                allResults.append("\n변형 ").append(index).append(": ").append(result.getText());
                if (normalized.contains("GHD18070916D7")) matched = true;
            }
            assertTrue(
                    "사진 인식값에 기대 각인이 없습니다." + allResults,
                    matched
            );
        } finally {
            recognizer.close();
            for (Bitmap variant : variants) {
                if (variant != bitmap && !variant.isRecycled()) variant.recycle();
            }
            bitmap.recycle();
        }
    }

    private List<Bitmap> createVariants(Bitmap original) {
        List<Bitmap> variants = new ArrayList<>();
        variants.add(original);
        int targetWidth = 1600;
        int targetHeight = Math.max(1, Math.round(
                original.getHeight() * (targetWidth / (float) original.getWidth())
        ));
        Bitmap scaled = Bitmap.createScaledBitmap(
                original,
                targetWidth,
                targetHeight,
                true
        );
        variants.add(scaled);
        variants.add(toHighContrast(scaled));
        variants.add(toBinary(scaled, 90));
        variants.add(toBinary(scaled, 120));
        variants.add(toBinary(scaled, 150));
        variants.add(toBinary(scaled, 180));
        variants.add(toBinary(scaled, 210));
        return variants;
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
}
