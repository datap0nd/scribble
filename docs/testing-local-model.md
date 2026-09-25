# Qwen for the local Office Test Bench

## Recommended on this PC: hosted Qwen3.8 27B

For the full Excel, PowerPoint and Outlook suite, configure Scribble's OpenAI-compatible provider with:

| Setting | Value |
| --- | --- |
| Endpoint | `https://openrouter.ai/api/v1` |
| Model | `qwen/qwen3.8-27b` |
| API key | Enter directly in Scribble Settings' API-key password field |

Use a dedicated key with a **$10 total spending limit**, no periodic limit reset, and no automatic account top-up. Creating or funding that key remains the owner's action; this guide does not purchase credits. OpenRouter documents a per-key USD limit and an optional reset interval in its [key settings API](https://openrouter.ai/docs/api/api-reference/api-keys/create-keys). Keep the key out of chat, scripts, reports and version control. Scribble stores a saved key using Windows encryption scoped to the current user; a permanent plaintext environment variable is not required for normal use.

At **2026-09-15 07:35 UTC**, OpenRouter's [model catalog API](https://openrouter.ai/api/v1/models) reported **$0.214 per million input tokens** and **$2.55 per million output tokens** for this model. These are dated catalog rates, not a guaranteed quote: the [model page](https://openrouter.ai/qwen/qwen3.8-27b) lists multiple providers and prices. The catalog advertises tool calling and image input. Actual Scribble compatibility, cost and output quality still require a run with the configured key.

Save Settings, use Scribble's endpoint check, then run a single case before the complete 16-case suite. Keep all test inputs synthetic. Public Scribble remains frozen at 2.0.91; install the tested development artifact separately as described in [release channels](release-channels.md).

## Local hosting measured on 2026-09-15

This owner's PC has an RTX 3070 with 8 GiB VRAM, 32 GB system RAM and a Ryzen 9800X3D. The official [Ollama Qwen3.8 27B](https://ollama.com/library/qwen3.8:27b) model was downloaded using portable Ollama 0.34.0: **27.3B Q4_K_M**, with its vision projector, totaling 17.74 GB on disk. No smaller model was substituted.

A **4096-token-context, non-thinking arithmetic probe** returned the correct 46,000 profit and 38.33% margin. It took 68 seconds including 44 seconds loading; generation reached about 4.2 tokens/second. Physical RAM fell from 11.58 GiB free to 0.78 GiB, with a lower intermediate reading of 0.51 GiB. This leaves insufficient headroom for Office. No tool-call or 16K-context probe was run, and the 4K result is **not Scribble validation**. The work PC's previous `fast` alias settings were not reproduced.

The model was unloaded immediately after the measurement. At setup completion the portable server remained idle on **127.0.0.1:11434 only**, with cloud features disabled; no startup task, firewall change or global PATH change was added. The weights remain in `%LOCALAPPDATA%\Scribble\LocalModels\ollama` for future use. Scribble settings were not changed by local runtime setup.

On this owner's checkout only, setup evidence and reproduction scripts are in `C:\Users\keeoh\Documents\ChatGPT\Scribble\artifacts\local-model-runtime`, including `README.md`, `findings.md`, `runtime.json`, `Start-LocalOllama.ps1` and `Test-LocalQwen.ps1`. These session artifacts are not shipped with Scribble or tracked in Git. They record the official archive checksum, exact model digest and raw timing results. Consult the [official Windows standalone instructions](https://docs.ollama.com/windows) for another computer; allow adequate RAM and VRAM before combining local inference with Office.
