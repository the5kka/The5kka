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

    @Test
    public void passwordHistorySearchAndTenMinuteResetWork() throws Exception {
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
            evaluate(webView,
                    "localStorage.removeItem('kccHdiLaserHistoryPasswordV1');"
                            + "localStorage.setItem('kccHdiLaserInspectionHistoryV1',JSON.stringify(["
                            + "{id:'2026-08-19T09:10:00.000Z-a',completedAtIso:'2026-08-19T09:10:00.000Z',"
                            + "startedAtIso:'2026-08-19T09:08:30.000Z',elapsedSeconds:90,judgment:'OK',"
                            + "management:'GHD1807D16',lot:'B7208901091601',manual:'7',"
                            + "expected:'GHD18070916D7',recognizedRaw:'GHD18070916D7',reason:'완전 일치'},"
                            + "{id:'2026-08-18T08:20:00.000Z-b',completedAtIso:'2026-08-18T08:20:00.000Z',"
                            + "startedAtIso:'2026-08-18T08:19:30.000Z',elapsedSeconds:30,judgment:'NG',"
                            + "management:'ABC1234E56',lot:'B7208901010801',manual:'A',"
                            + "expected:'ABC12340108EA',recognizedRaw:'ABC12340108EB',reason:'생성값과 불일치'}]));"
                            + "updateHistoryBadge();");

            assertEquals("5300", evaluate(webView, "getHistoryPassword();"));

            evaluate(webView, "document.getElementById('progressButton').click();");
            JSONObject locked = new JSONObject(evaluate(webView,
                    "JSON.stringify({passwordOpen:document.getElementById('historyPasswordDialog').open,"
                            + "historyOpen:document.getElementById('historyDialog').open})"));
            assertTrue("진행 현황은 먼저 비밀번호 창을 열어야 합니다.", locked.getBoolean("passwordOpen"));
            assertTrue("비밀번호 확인 전에는 작업이력을 열면 안 됩니다.", !locked.getBoolean("historyOpen"));

            evaluate(webView,
                    "document.getElementById('historyPasswordInput').value='1111';"
                            + "unlockHistory({preventDefault:function(){}});");
            assertTrue("틀린 비밀번호 경고가 표시되어야 합니다.",
                    evaluate(webView, "document.getElementById('historyPasswordMessage').textContent;")
                            .contains("맞지"));

            evaluate(webView,
                    "document.getElementById('historyPasswordInput').value='5300';"
                            + "unlockHistory({preventDefault:function(){}});");
            JSONObject opened = new JSONObject(evaluate(webView,
                    "JSON.stringify({open:document.getElementById('historyDialog').open,"
                            + "cards:document.querySelectorAll('.history-record-card').length,"
                            + "groups:document.querySelectorAll('.history-date-group').length})"));
            assertTrue("정상 비밀번호 입력 시 작업이력이 열려야 합니다.", opened.getBoolean("open"));
            assertEquals(2, opened.getInt("cards"));
            assertEquals(2, opened.getInt("groups"));

            evaluate(webView,
                    "document.getElementById('historyLotSearch').value='B7208901091601';renderHistory();");
            JSONObject filtered = new JSONObject(evaluate(webView,
                    "JSON.stringify({cards:document.querySelectorAll('.history-record-card').length,"
                            + "summary:document.getElementById('historyFilterSummary').textContent})"));
            assertEquals(1, filtered.getInt("cards"));
            assertTrue("LOT 검색 결과 건수가 표시되어야 합니다.", filtered.getString("summary").contains("1건"));

            evaluate(webView,
                    "openPasswordSettings('history');"
                            + "document.getElementById('masterPasswordInput').value='0160';"
                            + "document.getElementById('newHistoryPasswordInput').value='7788';"
                            + "document.getElementById('confirmHistoryPasswordInput').value='7788';"
                            + "changeHistoryPassword({preventDefault:function(){}});");
            assertEquals("7788", evaluate(webView, "getHistoryPassword();"));
            assertEquals("true", evaluate(webView, "document.getElementById('historyDialog').open;"));

            evaluate(webView,
                    "closeDialog(document.getElementById('historyDialog'));openHistoryPasswordDialog();"
                            + "document.getElementById('historyPasswordInput').value='7788';"
                            + "unlockHistory({preventDefault:function(){}});"
                            + "closeDialog(document.getElementById('historyDialog'));");

            fillValidWork(webView);
            JSONObject timer = new JSONObject(evaluate(webView,
                    "JSON.stringify({remaining:autoResetDeadline-Date.now(),"
                            + "label:document.getElementById('autoResetStatus').textContent,"
                            + "result:document.getElementById('resultValue').textContent})"));
            assertTrue("자동 초기화 시간은 약 10분이어야 합니다.",
                    timer.getLong("remaining") >= 590000 && timer.getLong("remaining") <= 601000);
            assertTrue("10분 카운트다운 안내가 표시되어야 합니다.", timer.getString("label").contains("10:00"));
            assertEquals("GHD18070916D7", timer.getString("result"));

            evaluate(webView, "autoResetDeadline=Date.now()-1;handleAutoResetTick();");
            JSONObject reset = new JSONObject(evaluate(webView,
                    "JSON.stringify({management:document.getElementById('managementNo').value,"
                            + "result:currentResult,"
                            + "history:JSON.parse(localStorage.getItem('kccHdiLaserInspectionHistoryV1')).length})"));
            assertEquals("", reset.getString("management"));
            assertEquals("", reset.getString("result"));
            assertEquals("자동 초기화가 작업이력을 지우면 안 됩니다.", 2, reset.getInt("history"));
        } finally {
            evaluate(webView,
                    "localStorage.removeItem('kccHdiLaserInspectionHistoryV1');"
                            + "localStorage.removeItem('kccHdiLaserHistoryPasswordV1');");
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
        for (int attempt = 0; attempt < 50; attempt++) {
            String result = evaluate(webView,
                    "document.readyState==='complete'"
                            + "&&typeof showJudgmentDialog==='function'"
                            + "&&document.getElementById('progressButton')!==null;");
            if ("true".equals(result)) return;
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
