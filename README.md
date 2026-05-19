# AutoTOS

AI-assisted Table of Specification (TOS) and exam item generator. Built as a thesis project to help educators automatically generate exam questions aligned with Bloom's taxonomy.

## Stack

- **Frontend:** Flask + Jinja templates
- **Orchestration:** FastAPI service layer
- **AI:** Ollama serving QLoRA fine-tuned Qwen3-4B-Instruct (Q4 quantization)
- **Database:** MySQL 8
- **Deployment:** Docker Compose on Hostinger VPS

## Architecture

[Browser]
|
v
[Flask web (:5000)] --> [FastAPI ai (:8000)] --> [Ollama (:11434)]
|
v
[MySQL db (:3306)]

## Features

- Generate Multiple Choice, True/False, and Open-Ended questions
- Bloom's taxonomy distribution (default: Familiarization 50% / Integration 30% / Creation 20%)
- Custom Bloom percentages supported
- Background extraction of learning materials (PDF, DOCX, PPTX, TXT)
- Persistent learning materials with in-app review
- Test Distribution table showing item ranges per test section
- DOCX export of complete TOS + exam
- Multi-user with Google OAuth login

## Local development

1. Clone this repo
2. Copy `.env.example` to `.env` and fill in your values
3. Download the GGUF model and place at `models/autotos-q4_k_m.gguf`
4. Run `docker compose up -d`
5. Visit http://localhost:5000

## Deployment

Deployed on a Hostinger VPS with Docker Compose. CPU inference via Ollama (Q4 quantization).

## License

Private — academic thesis project.
