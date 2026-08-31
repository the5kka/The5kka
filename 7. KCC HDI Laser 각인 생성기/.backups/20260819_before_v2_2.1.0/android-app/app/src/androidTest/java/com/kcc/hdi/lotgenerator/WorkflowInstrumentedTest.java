package com.kcc.hdi.lotgenerator;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import android.app.Instrumentation;
import android.content.Intent;
import android.webkit.WebView;

import androidx.test.ext.junit.runners.AndroidJUnit4;
import androidx.test.platform.app.InstrumentationRegistry;

import org.json.JSONObject;
import org.junit.Test;
import org.junit.runner.RunWith;

import java.lang.reflect.Field;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

@RunWith(AndroidJUnit4.class)
public class WorkflowInstrumentedTest {
    @Test
    public void okAndNgDialogsSaveHistoryAndClearCurrentWork() throws Exception {
        Instrumentation instrumentation = InstrumentationRegistry.getInstrumentation();
        Intent intent = instrumentation.getTargetContext()
                .getPackageManager()
                .getLaunchIntentForPackage(instrumentation.getTargetContext().getPackageName());
        assertTrue("V2 실행 Intent를 찾지 못했습니다.", intent != null);
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_CLEAR_TASK);

        MainActivity activity = (MainActivity) instrumentation.startActivitySync(intent);
        WebView webView = getWebView(activity);
        waitForPage(webView);

        try {
            evaluate(webView, "localStorage.removeItem('kccHdiLaserInspectionHistoryV1');");
            fillValidWork(webView);
            evaluate(webView,
                    "showJudgmentDialog('OK','GHD18070916D7 31','GHD18070916D731','완전 일치');");

            JSONObject okDialog = new JSONObject(evaluate(webView,
                    "JSON.stringify({title:document.getElementById('judgmentTitle').textContent,"
                            + "retryHidden:document.getElementById('retryPhotoButton').classList.contains('is-hidden')})"));
            assertEquals("초도품 각인 대조 OK", okDialog.getString("title"));
            assertTrue("OK에서는 다시 찍기 버튼이 숨겨져야 합니다.", okDialog.getBoolean("retryHidden"));

            evaluate(webView, "document.getElementById('confirmJudgmentButton').click();");
            JSONObject afterOk = new JSONObject(evaluate(webView,
                    "JSON.stringify({count:JSON.parse(localStorage.getItem('kccHdiLaserInspectionHistoryV1')).length,"
                            + "management:document.getElementById('managementNo').value})"));
            assertEquals(1, afterOk.getInt("count"));
            assertEquals("", afterOk.getString("management"));

            fillValidWork(webView);
            evaluate(webView,
                    "showJudgmentDialog('NG','GHD1807091607 31','GHD180709160731','생성값과 불일치');");
            JSONObject ngDialog = new JSONObject(evaluate(webView,
                    "JSON.stringify({title:document.getElementById('judgmentTitle').textContent,"
                            + "retryHidden:document.getElementById('retryPhotoButton').classList.contains('is-hidden'),"
                            + "management:document.getElementById('managementNo').value})"));
            assertEquals("조장 확인 필요!!", ngDialog.getString("title"));
            assertTrue("NG에서는 다시 찍기 버튼이 보여야 합니다.", !ngDialog.getBoolean("retryHidden"));
            assertEquals("GHD1807D16", ngDialog.getString("management"));

            evaluate(webView, "document.getElementById('confirmJudgmentButton').click();");
            JSONObject afterNg = new JSONObject(evaluate(webView,
                    "JSON.stringify({history:JSON.parse(localStorage.getItem('kccHdiLaserInspectionHistoryV1')),"
                            + "management:document.getElementById('managementNo').value})"));
            assertEquals(2, afterNg.getJSONArray("history").length());
            assertEquals("NG", afterNg.getJSONArray("history").getJSONObject(0).getString("judgment"));
            assertEquals("", afterNg.getString("management"));
        } finally {
            evaluate(webView, "localStorage.removeItem('kccHdiLaserInspectionHistoryV1');");
            instrumentation.runOnMainSync(activity::finish);
        }
    }

    private void fillValidWork(WebView webView) throws Exception {
        evaluate(webView,
                "window.onAndroidManagementScanned('GHD1807D16');"
                        + "window.onAndroidLotScanned('B7208901091601');"
                        + "document.getElementById('manualValue').value='7';"
                        + "updateResult();");
    }

    private WebView getWebView(MainActivity activity) throws Exception {
        Field field = MainActivity.class.getDeclaredField("webView");
        field.setAccessible(true);
        return (WebView) field.get(activity);
    }

    private void waitForPage(WebView webView) throws Exception {
        for (int attempt = 0; attempt < 20; attempt++) {
            String result = evaluate(webView, "document.readyState");
            if ("complete".equals(result)) return;
            Thread.sleep(100);
        }
        throw new AssertionError("앱 화면 로딩이 완료되지 않았습니다.");
    }

    private String evaluate(WebView webView, String script) throws Exception {
        CountDownLatch latch = new CountDownLatch(1);
        AtomicReference<String> result = new AtomicReference<>("");
        InstrumentationRegistry.getInstrumentation().runOnMainSync(() ->
                webView.evaluateJavascript(script, value -> {
                    result.set(value == null ? "" : value);
                    latch.countDown();
                })
        );
        assertTrue("JavaScript 검사 시간이 초과되었습니다.", latch.await(5, TimeUnit.SECONDS));
        String value = result.get();
        if (value.startsWith("\"") && value.endsWith("\"")) {
            return new org.json.JSONTokener(value).nextValue().toString();
        }
        return value;
    }
}
