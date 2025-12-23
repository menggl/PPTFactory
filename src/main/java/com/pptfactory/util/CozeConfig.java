package com.pptfactory.util;

/**
 * Coze 配置常量（从环境变量读取）。
 * 在类加载时读取一次，供工具类使用。
 * 注意：本仓库后续生成的代码注释请采用中文。
 */
public final class CozeConfig {
    // Coze workflow 的默认 endpoint（默认使用 /v1/workflow/run，可由 COZE_ENDPOINT 覆盖）
    public static final String ENDPOINT = "https://api.coze.cn/v1/workflow/run";

    // 用于 Authorization: Bearer <token> 的 token（可由 COZE_TOKEN 或 COZE_API_KEY 环境变量覆盖）
    // 缺省值已设置为当前运行常用 token（可在 CI/环境中通过环境变量覆盖以避免在代码中存储明文）。
    // 当需要强制使用代码中配置的默认 token 时，请使用 GenerateImagesViaCozeUtil 的 --use-config-token 选项。
    public static final String DEFAULT_TOKEN = "pat_CNSVEaa3BbRFHThYjyqsp3wvR7ZogthTpG0rK5jj200dO32ETYhK4lTdTB5DhxXz";
    public static final String TOKEN =  DEFAULT_TOKEN;

    // 专用于生成图片的 workflow id（可由 COZE_WORKFLOW_ID 覆盖）
    public static final String IMAGE_WORKFLOW_ID = "7582876497963647027";

    // 兼容字段：旧代码仍可通过 WORKFLOW_ID 访问（指向 IMAGE_WORKFLOW_ID）
    public static final String WORKFLOW_ID = IMAGE_WORKFLOW_ID;

    // parameters 的默认 type（可选）
    public static final int TYPE =  0;

    // 重试 / 超时 / 回退 设置（为避免重复消费 token，默认重试次数设置为 1）
    public static final int MAX_RETRIES = 1;
    public static final int TIMEOUT_SECONDS = 120;
    // 回退基准毫秒（sleep = RETRY_BACKOFF_MS * attempt）
    public static final int RETRY_BACKOFF_MS = 500;

    // 每次请求后的小延迟以避免突发请求（毫秒）
    public static final int REQUEST_SLEEP_MS =  300;

    // 授权头相关常量（header 名称与前缀）
    public static final String AUTH_HEADER = "Authorization";
    public static final String AUTH_PREFIX = "Bearer ";

    private static int parseInt(String s, int fallback) {
        try { return Integer.parseInt(s); } catch (Exception e) { return fallback; }
    }

    private CozeConfig() {}
}
