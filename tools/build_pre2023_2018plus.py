from pathlib import Path

import build_pre2023_supplement as source


source.MIN_YEAR = 2018
source.DEDUPLICATE_AGAINST_BASE = False
source.OUTPUT_DOCX = Path(__file__).resolve().parents[1] / "readme_2018_to_2022.docx"
source.DOCUMENT_TITLE = "Multi-Turn Speech Dialogue: 2018-2022 Supplement"
source.INTRO_TEXT = (
    "Papers retained from readme_pre_2023.docx after removing entries "
    "published before 2018."
)
source.CUTOFF_TEXT = "2018-2022"
source.COUNT_LABEL = "Retained papers"
source.SUMMARY_THIRD_LABEL = "Source filter"
source.SUMMARY_THIRD_TEXT = "60 pre-2018 entries removed"
source.EMPTY_SECTION_TEXT = "No paper from 2018 through 2022 is listed in this subsection."
source.HEADER_TEXT = "Multi-Turn Speech Dialogue | 2018-2022 Supplement"
source.SUBJECT_TEXT = "Curated papers published from 2018 through 2022"
source.RECLASSIFICATIONS = {
    "A Survey of Available Corpora For Building Data-Driven Dialogue Systems: The Journal Version": ("General Surveys", None),
    "Improving End-of-Turn Detection in Spoken Dialogues by Detecting Speaker Intentions as a Secondary Task": ("Models", "Interaction Mode: Turn-taking"),
    "Neural Dialogue Context Online End-of-Turn Detection": ("Models", "Interaction Mode: Turn-taking"),
    "Prediction of Turn-Taking Using Multitask Learning with Prediction of Backchannels and Fillers": ("Models", "Interaction Mode: Turn-taking"),
    "An Incremental Turn-Taking Model for Task-Oriented Dialog Systems": ("Models", "Interaction Mode: Half-duplex / Controlled Barge-in"),
    "Turn-Taking Prediction Based on Detection of Transition Relevance Place": ("Models", "Interaction Mode: Turn-taking"),
    "TurnGPT: A Transformer-Based Language Model for Predicting Turn-Taking in Spoken Dialog": ("Models", "Interaction Mode: Turn-taking"),
    "Projection of Turn Completion in Incremental Spoken Dialogue Systems": ("Models", "Interaction Mode: Turn-taking"),
    "Timing Generating Networks: Neural Network Based Precise Turn-Taking Timing Prediction in Multiparty Conversation": ("Models", "Interaction Mode: Turn-taking"),
    "Gated Multimodal Fusion with Contrastive Learning for Turn-Taking Prediction in Human-Robot Dialogue": ("Models", "Interaction Mode: Turn-taking"),
    "Response Timing Estimation for Spoken Dialog System Using Dialog Act Estimation": ("Models", "Interaction Mode: Turn-taking"),
    "Turn-Taking Prediction for Natural Conversational Speech": ("Models", "Interaction Mode: Turn-taking"),
    "Voice Activity Projection: Self-Supervised Learning of Turn-Taking Events": ("Models", "Interaction Mode: Turn-taking"),
    "Oh, Jeez! or Uh-Huh? A Listener-Aware Backchannel Predictor on ASR Transcriptions": ("Models", "Interaction Mode: Half-duplex / Controlled Barge-in"),
    "BPM_MT: Enhanced Backchannel Prediction Model Using Multi-Task Learning": ("Models", "Interaction Mode: Half-duplex / Controlled Barge-in"),
    "Contextual Acoustic Barge-In Classification for Spoken Dialog Systems": ("Models", "Interaction Mode: Half-duplex / Controlled Barge-in"),
}


if __name__ == "__main__":
    source.build()
