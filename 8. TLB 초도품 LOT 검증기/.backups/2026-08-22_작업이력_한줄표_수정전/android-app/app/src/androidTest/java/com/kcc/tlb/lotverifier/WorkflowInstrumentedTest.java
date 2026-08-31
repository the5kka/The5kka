package com.kcc.tlb.lotverifier;

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
    public void validatesNineCharacterLotsAndCompletesOkNgHistory() throws Exception {
        Instrumentation instrumentation = InstrumentationRegistry.getInstrumentation();
        Intent intent = instrumentation.getTargetContext().getPackageManager()
                .getLaunchIntentForPackage(instrumentation.getTargetContext().getPackageName());
        assertTrue("TLB 실행 Intent를 찾지 못했습니다.", intent != null);
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_CLEAR_TASK);

        MainActivity activity = (MainActivity) instrumentation.startActivitySync(intent);
        WebView webView = getWebView(activity);
        waitForPage(webView);
        backupStorage(webView);

        try {
            evaluate(webView,
                    "localStorage.removeItem('tlbLotInspectionHistoryV1');"
                            + "localStorage.removeItem('tlbLotHistoryPasswordV1');"
                            + "localStorage.removeItem('tlbLotAutoResetMinutesV1');");

            evaluate(webView, "window.onAndroidLotScanned('362109B09');");
            JSONObject first = new JSONObject(evaluate(webView,
                    "JSON.stringify({lot:document.getElementById('lotNo').value,"
                            + "result:document.getElementById('resultValue').textContent,"
                            + "remaining:autoResetDeadline-Date.now(),"
                            + "label:document.getElementById('autoResetStatus').textContent})"));
            assertEquals("362109B09", first.getString("lot"));
            assertEquals("362109B09", first.getString("result"));
            assertTrue("TLB 기본 자동 초기화는 약 10분이어야 합니다.",
                    first.getLong("remaining") >= 590000 && first.getLong("remaining") <= 601000);
            assertTrue(first.getString("label").contains("10:00"));

            evaluate(webView, "window.onAndroidLotScanned('362109B0');");
            JSONObject invalid = new JSONObject(evaluate(webView,
                    "JSON.stringify({current:currentLot,error:document.getElementById('lotMessage').textContent})"));
            assertEquals("", invalid.getString("current"));
            assertTrue("8자리 LOT는 차단해야 합니다.", invalid.getString("error").contains("9자리"));

            evaluate(webView,
                    "window.onAndroidLotScanned('S16189007');"
                            + "showJudgmentDialog('OK','S16189007','S16189007','카드 LOT와 사진 LOT가 일치합니다.');"
                            + "document.getElementById('confirmJudgmentButton').click();");
            JSONObject afterOk = new JSONObject(evaluate(webView,
                    "JSON.stringify({count:JSON.parse(localStorage.getItem('tlbLotInspectionHistoryV1')).length,"
                            + "lot:document.getElementById('lotNo').value})"));
            assertEquals(1, afterOk.getInt("count"));
            assertEquals("", afterOk.getString("lot"));

            evaluate(webView,
                    "window.onAndroidLotScanned('362109B09');"
                            + "showJudgmentDialog('NG','362109B08','362109B08','카드 LOT와 사진 인식값이 일치하지 않습니다.');");
            JSONObject ngDialog = new JSONObject(evaluate(webView,
                    "JSON.stringify({title:document.getElementById('judgmentTitle').textContent,"
                            + "retryVisible:!document.getElementById('retryPhotoButton').classList.contains('is-hidden')})"));
            assertEquals("조장 확인 필요!!", ngDialog.getString("title"));
            assertTrue("NG에서는 다시 찍기가 보여야 합니다.", ngDialog.getBoolean("retryVisible"));
            evaluate(webView, "document.getElementById('confirmJudgmentButton').click();");
            assertEquals("2", evaluate(webView,
                    "JSON.parse(localStorage.getItem('tlbLotInspectionHistoryV1')).length;"));

            assertEquals("5300", evaluate(webView, "getHistoryPassword();"));
            evaluate(webView,
                    "document.getElementById('autoResetMinutesSelect').value='5';"
                            + "saveAppSettings({preventDefault:function(){}});"
                            + "window.onAndroidLotScanned('S16189007');");
            long fiveMinuteRemaining = Long.parseLong(evaluate(webView, "Math.round(autoResetDeadline-Date.now());"));
            assertTrue("설정에서 5분으로 변경할 수 있어야 합니다.",
                    fiveMinuteRemaining >= 290000 && fiveMinuteRemaining <= 301000);

            evaluate(webView, "autoResetDeadline=Date.now()-1;handleAutoResetTick();");
            JSONObject reset = new JSONObject(evaluate(webView,
                    "JSON.stringify({lot:document.getElementById('lotNo').value,"
                            + "history:JSON.parse(localStorage.getItem('tlbLotInspectionHistoryV1')).length})"));
            assertEquals("", reset.getString("lot"));
            assertEquals("자동 초기화가 작업이력을 지우면 안 됩니다.", 2, reset.getInt("history"));
        } finally {
            restoreStorage(webView);
            instrumentation.runOnMainSync(activity::finish);
        }
    }

    private void backupStorage(WebView webView) throws Exception {
        evaluate(webView,
                "window.__tlbTestStorageBackup={"
                        + "history:localStorage.getItem('tlbLotInspectionHistoryV1'),"
                        + "password:localStorage.getItem('tlbLotHistoryPasswordV1'),"
                        + "reset:localStorage.getItem('tlbLotAutoResetMinutesV1')};");
    }

    private void restoreStorage(WebView webView) throws Exception {
        evaluate(webView,
                "(function(){var b=window.__tlbTestStorageBackup;if(!b)return;"
                        + "[['tlbLotInspectionHistoryV1',b.history],"
                        + "['tlbLotHistoryPasswordV1',b.password],"
                        + "['tlbLotAutoResetMinutesV1',b.reset]].forEach(function(p){"
                        + "if(p[1]===null)localStorage.removeItem(p[0]);else localStorage.setItem(p[0],p[1]);});"
                        + "delete window.__tlbTestStorageBackup;})();");
    }

    private WebView getWebView(MainActivity activity) throws Exception {
        Field field = MainActivity.class.getDeclaredField("webView");
        field.setAccessible(true);
        return (WebView) field.get(activity);
    }

    private void waitForPage(WebView webView) throws Exception {
        for (int attempt = 0; attempt < 50; attempt++) {
            String result = evaluate(webView,
                    "document.readyState==='complete'"
                            + "&&typeof updateLot==='function'"
                            + "&&document.getElementById('progressButton')!==null;");
            if ("true".equals(result)) return;
            Thread.sleep(100);
        }
        throw new AssertionError("TLB 앱 화면 로딩이 완료되지 않았습니다.");
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
