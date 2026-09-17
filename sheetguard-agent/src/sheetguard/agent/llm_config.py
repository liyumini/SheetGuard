"""Provider 无关的 LLM 配置模块，用于构建 OpenAI 兼容的模型。

通过环境变量 OPENAI_MODEL / OPENAI_BASE_URL / OPENAI_API_KEY 创建模型，
换 provider 或代理时只需修改 .env，不需要改代码。
"""
from __future__ import annotations
import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.language_models import BaseChatModel

# 常见的占位符模型名（web 状态判定与 build_model 校验共用）。
PLACEHOLDER_MODELS = {"your-model-name", "your_model_name"}
# 常见的占位符 API key（与 CLI 的配置门判定保持一致）。
PLACEHOLDER_KEYS = {"", "your-api-key-here", "sk-...", "sk-placeholder"}


def model_configured() -> bool:
    """key/model 均非占位符时视为已配置（不打网络请求）。"""
    load_dotenv()
    model = (os.getenv("OPENAI_MODEL") or "").strip()
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    return (
        bool(model)
        and model.casefold() not in PLACEHOLDER_MODELS
        and key.lower() not in PLACEHOLDER_KEYS
    )


def build_model() -> BaseChatModel:
    """从环境变量构建一个 OpenAI 兼容的 ChatOpenAI 模型。

    通过 OPENAI_BASE_URL 和 OPENAI_MODEL 等环境变量配置模型，
    这两个变量故意做成 provider 无关，因此更换 OpenAI 兼容的代理时，
    只需要修改 .env 文件即可。
    """
    load_dotenv()  # 从项目根目录的 .env 文件加载环境变量
    model = (os.getenv("OPENAI_MODEL") or "").strip()  # 读取模型名，去掉首尾空格
    placeholder_models = PLACEHOLDER_MODELS  # 常见的占位符模型名
    if not model or model.casefold() in placeholder_models:
        # 模型名为空或仍是占位符时直接报错，避免用假模型名请求 API
        raise ValueError(
            "OPENAI_MODEL must be configured with a real provider model name; "
            "replace the placeholder value"
        )
    return ChatOpenAI(
        model=model,  # 模型名
        base_url=os.getenv("OPENAI_BASE_URL") or None,  # API 代理地址，未配置则用默认
        api_key=os.getenv("OPENAI_API_KEY", "sk-placeholder"),  # API 密钥
        temperature=0.0,  # 温度 0，输出更稳定、可复现
    )
