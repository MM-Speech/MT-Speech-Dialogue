# MT-Speech-Dialogue

> A curated, chronological reading list for **Multi-Turn Speech Dialogue**. The organization follows the paper-list style of [LLMSurvey](https://github.com/RUCAIBox/LLMSurvey): title, authors, venue/year, paper link, and a reusable citation key.
>
> Scope: systems that preserve and use dialogue history across turns, including spoken task-oriented dialogue, speech-to-speech conversational agents, and full-duplex spoken dialogue. Papers may appear in more than one category when they introduce both a resource and an evaluation or training method.

## Paper List

### Surveys

1. **"Spoken Dialogue Technology: Enabling the Conversational User Interface"**. *Michael McTear.* ACM Computing Surveys 2002. [[Paper](https://doi.org/10.1145/505282.505285)] [`McTear2002`]
2. **"A Survey on Dialogue Systems: Recent Advances and New Frontiers"**. *Hongshen Chen et al.* SIGKDD Explorations 2017. [[Paper](https://arxiv.org/abs/1711.01731)] [`Chen2017`]
3. **"A Survey on Speech Large Language Models"**. *Jing Peng et al.* arXiv 2024. [[Paper](https://arxiv.org/abs/2410.18908)] [`Peng2024`]
4. **"A Survey on Recent Advances in LLM-Based Multi-turn Dialogue Systems"**. *Zihao Yi et al.* arXiv 2024. [[Paper](https://arxiv.org/abs/2402.18013)] [`Yi2024`]

### Datasets

1. **"SWITCHBOARD: Telephone Speech Corpus for Research and Development"**. *John J. Godfrey et al.* ICASSP 1992. [[Paper](https://doi.org/10.1109/ICASSP.1992.225858)] [`Godfrey1992`]
2. **"The ICSI Meeting Corpus"**. *Adam Janin et al.* ICASSP 2003. [[Paper](https://doi.org/10.1109/ICASSP.2003.1202314)] [`Janin2003`]
3. **"The Fisher Corpus: A Resource for the Next Generations of Speech-to-Text"**. *Christopher Cieri et al.* LREC 2004. [[Paper](https://aclanthology.org/L04-1300/)] [`Cieri2004`]
4. **"The Second Dialog State Tracking Challenge"**. *Matthew Henderson et al.* SIGDIAL 2014. [[Paper](https://aclanthology.org/W14-4337/)] [`Henderson2014`]
5. **"SpokenWOZ: A Large-Scale Speech-Text Benchmark for Spoken Task-Oriented Dialogue Agents"**. *Shuzheng Si et al.* NeurIPS Datasets and Benchmarks 2023. [[Paper](https://arxiv.org/abs/2305.13040)] [`Si2023`]
6. **"RealTalk-CN: A Realistic Chinese Speech-Text Dialogue Benchmark With Cross-Modal Interaction Analysis"**. *Enzhi Wang et al.* arXiv 2025. [[Paper](https://arxiv.org/abs/2508.10015)] [`Wang2025`]

### Evaluation

1. **"PARADISE: A Framework for Evaluating Spoken Dialogue Agents"**. *Marilyn A. Walker et al.* ACL/EACL 1997. [[Paper](https://aclanthology.org/P97-1035/)] [`Walker1997`]
2. **"The Second Dialog State Tracking Challenge"**. *Matthew Henderson et al.* SIGDIAL 2014. Shared spoken-dialogue test set, metrics, and evaluation scripts. [[Paper](https://aclanthology.org/W14-4337/)] [`Henderson2014`]
3. **"The Dialog State Tracking Challenge Series: A Review"**. *Jason D. Williams et al.* Dialogue & Discourse 2016. [[Paper](https://aclanthology.org/2016.dnd-7.5/)] [`Williams2016`]
4. **"SpokenWOZ: A Large-Scale Speech-Text Benchmark for Spoken Task-Oriented Dialogue Agents"**. *Shuzheng Si et al.* NeurIPS Datasets and Benchmarks 2023. Evaluates cross-turn slots, reasoning slots, ASR noise, DST, and end-to-end task success. [[Paper](https://arxiv.org/abs/2305.13040)] [`Si2023`]
5. **"VoiceBench: Benchmarking LLM-Based Voice Assistants"**. *Yiming Chen et al.* arXiv 2024. [[Paper](https://arxiv.org/abs/2410.17196)] [`Chen2024VoiceBench`]

### Models

1. **"Generative Spoken Dialogue Language Modeling"**. *Tu Anh Nguyen et al.* TACL 2023 (first released 2022). dGSLM directly models two-channel spoken dialogue without text. [[Paper](https://aclanthology.org/2023.tacl-1.15/)] [`Nguyen2023`]
2. **"SpeechGPT: Empowering Large Language Models with Intrinsic Cross-Modal Conversational Abilities"**. *Dong Zhang et al.* arXiv 2023. [[Paper](https://arxiv.org/abs/2305.11000)] [`Zhang2023SpeechGPT`]
3. **"Qwen2-Audio Technical Report"**. *Yunfei Chu et al.* arXiv 2024. Voice-chat and audio-analysis modes for audio-language interaction. [[Paper](https://arxiv.org/abs/2407.10759)] [`Chu2024`]
4. **"Mini-Omni: Language Models Can Hear, Talk While Thinking in Streaming"**. *Zhifei Xie and Changqiao Wu.* arXiv 2024. [[Paper](https://arxiv.org/abs/2408.16725)] [`Xie2024`]
5. **"LLaMA-Omni: Seamless Speech Interaction with Large Language Models"**. *Qingkai Fang et al.* arXiv 2024. [[Paper](https://arxiv.org/abs/2409.06666)] [`Fang2024`]
6. **"Moshi: a speech-text foundation model for real-time dialogue"**. *Alexandre Defossez et al.* arXiv 2024. A real-time full-duplex speech-to-speech dialogue model. [[Paper](https://arxiv.org/abs/2410.00037)] [`Defossez2024`]

### Training Methods

1. **"POMDP-based Statistical Spoken Dialog Systems: A Review"**. *Steve Young et al.* Proceedings of the IEEE 2013. Belief-state tracking and reinforcement learning for multi-turn spoken dialogue policy optimization. [[Paper](https://doi.org/10.1109/JPROC.2012.2225812)] [`Young2013`]
2. **"Generative Spoken Dialogue Language Modeling"**. *Tu Anh Nguyen et al.* TACL 2023 (first released 2022). Self-supervised training on two-channel conversational audio using discrete speech units and cross-attention. [[Paper](https://aclanthology.org/2023.tacl-1.15/)] [`Nguyen2023`]
3. **"SpeechGPT: Empowering Large Language Models with Intrinsic Cross-Modal Conversational Abilities"**. *Dong Zhang et al.* arXiv 2023. Three stages: modality adaptation pre-training, cross-modal instruction tuning, and chain-of-modality instruction tuning. [[Paper](https://arxiv.org/abs/2305.11000)] [`Zhang2023SpeechGPT`]
4. **"BLSP: Bootstrapping Language-Speech Pre-training via Behavior Alignment of Continuation Writing"**. *Dong Zhang et al.* arXiv 2023. Aligns speech-conditioned continuations with frozen LLM behavior before downstream speech instruction tuning. [[Paper](https://arxiv.org/abs/2309.00916)] [`Zhang2023BLSP`]
5. **"Mini-Omni: Language Models Can Hear, Talk While Thinking in Streaming"**. *Zhifei Xie and Changqiao Wu.* arXiv 2024. Text-instructed speech generation and batch-parallel decoding for streaming spoken interaction. [[Paper](https://arxiv.org/abs/2408.16725)] [`Xie2024`]
6. **"Moshi: a speech-text foundation model for real-time dialogue"**. *Alexandre Defossez et al.* arXiv 2024. Parallel user/assistant audio streams, hierarchical audio-token prediction, and an inner-monologue text prefix for low-latency full-duplex dialogue. [[Paper](https://arxiv.org/abs/2410.00037)] [`Defossez2024`]

## Citation Keys

Use the identifier displayed after each paper in LaTeX-style references, for example: `\cite{Nguyen2023}`. The keys below are intentionally stable and match the labels in the paper list.

| Key | Reference |
| --- | --- |
| `McTear2002` | McTear, *Spoken Dialogue Technology*, ACM CS 2002. |
| `Chen2017` | Chen et al., *A Survey on Dialogue Systems*, SIGKDD Explorations 2017. |
| `Peng2024` | Peng et al., *A Survey on Speech Large Language Models*, arXiv 2024. |
| `Yi2024` | Yi et al., *A Survey on Recent Advances in LLM-Based Multi-turn Dialogue Systems*, arXiv 2024. |
| `Godfrey1992` | Godfrey et al., *SWITCHBOARD*, ICASSP 1992. |
| `Janin2003` | Janin et al., *The ICSI Meeting Corpus*, ICASSP 2003. |
| `Cieri2004` | Cieri et al., *The Fisher Corpus*, LREC 2004. |
| `Henderson2014` | Henderson et al., *The Second Dialog State Tracking Challenge*, SIGDIAL 2014. |
| `Si2023` | Si et al., *SpokenWOZ*, NeurIPS Datasets and Benchmarks 2023. |
| `Wang2025` | Wang et al., *RealTalk-CN*, arXiv 2025. |
| `Walker1997` | Walker et al., *PARADISE*, ACL/EACL 1997. |
| `Williams2016` | Williams et al., *The Dialog State Tracking Challenge Series: A Review*, Dialogue & Discourse 2016. |
| `Chen2024VoiceBench` | Chen et al., *VoiceBench*, arXiv 2024. |
| `Nguyen2023` | Nguyen et al., *Generative Spoken Dialogue Language Modeling*, TACL 2023. |
| `Zhang2023SpeechGPT` | Zhang et al., *SpeechGPT*, arXiv 2023. |
| `Zhang2023BLSP` | Zhang et al., *BLSP*, arXiv 2023. |
| `Chu2024` | Chu et al., *Qwen2-Audio Technical Report*, arXiv 2024. |
| `Xie2024` | Xie and Wu, *Mini-Omni*, arXiv 2024. |
| `Fang2024` | Fang et al., *LLaMA-Omni*, arXiv 2024. |
| `Defossez2024` | Defossez et al., *Moshi*, arXiv 2024. |
| `Young2013` | Young et al., *POMDP-based Statistical Spoken Dialog Systems*, Proceedings of the IEEE 2013. |
