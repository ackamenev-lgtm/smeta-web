import unittest

from bitrix_bot.answer import (
    classify,
    comment_body,
    render,
    select_pending,
    should_answer,
)


class ClassifyTests(unittest.TestCase):
    def test_help(self):
        self.assertEqual(classify("привет").kind, "help")

    def test_stage_uses_number_after_word(self):
        intent = classify("Какая стадия сделки 1930, закупка 4660147")
        self.assertEqual(intent.kind, "stage")
        self.assertEqual(intent.query, "1930")

    def test_documents(self):
        intent = classify("Где документы 4660147")
        self.assertEqual(intent.kind, "docs")
        self.assertEqual(intent.query, "4660147")

    def test_calc_lot(self):
        intent = classify("Посчитай лот 390-448")
        self.assertEqual(intent.kind, "calc")
        self.assertEqual(intent.query, "390-448")

    def test_bare_number_is_full_card(self):
        intent = classify("4660147")
        self.assertEqual(intent.kind, "all")
        self.assertEqual(intent.query, "4660147")


class AddressTests(unittest.TestCase):
    def test_personal_dialog(self):
        self.assertTrue(should_answer("стадия 1", "12", "chat9"))

    def test_bot_chat(self):
        self.assertTrue(should_answer("ок", "chat9", "chat9"))

    def test_general_chat_needs_mention(self):
        self.assertFalse(should_answer("доброе утро", "chat2", "chat9"))
        self.assertTrue(should_answer("[USER=316 REPLACE]AI Агент[/USER] стадия", "chat2", "chat9"))


class PendingTests(unittest.TestCase):
    def test_ignores_own_and_system_and_takes_latest_after_own(self):
        messages = [
            {"id": 1, "author_id": 5, "text": "старое"},
            {"id": 2, "author_id": 316, "text": "уже ответил"},
            {"id": 3, "author_id": 0, "text": "система"},
            {"id": 4, "author_id": 8, "text": "стадия 1930"},
            {"id": 5, "author_id": 8, "text": "где документы 1930"},
        ]
        pending = select_pending(messages)
        self.assertEqual(pending["id"], 5)

    def test_nothing_after_own_reply(self):
        messages = [
            {"id": 4, "author_id": 8, "text": "стадия 1930"},
            {"id": 5, "author_id": 316, "text": "ответ"},
        ]
        self.assertIsNone(select_pending(messages))


class RenderTests(unittest.TestCase):
    def setUp(self):
        self.deal = {
            "ID": "1930",
            "TITLE": "Реконструкция фуд-корта",
            "STAGE_ID": "C6:UC_0AH6PM",
            "UF_CRM_TP_11": "4660147",
            "UF_CRM_TP_10": "https://www.b2b-center.ru/tender-4660147/",
        }
        self.comment = (
            "Документы не скачаны.<br>"
            "Папка закупки 4660147: "
            "https://rps-official.bitrix24.ru/bitrix/tools/disk/focus.php?folderId=51132&action=openFolderList&ncc=1<br>"
            "конкурсная документация: пусто"
        )

    def test_stage(self):
        text = render(classify("стадия 1930"), self.deal, self.comment)
        self.assertIn("ИИ не может скачать документы", text)
        self.assertNotIn("focus.php", text)

    def test_docs_include_folder(self):
        text = render(classify("где документы 4660147"), self.deal, self.comment)
        self.assertIn("folderId=51132", text)
        self.assertIn("конкурсная документация: пусто", text)
        self.assertIn("b2b-center.ru", text)

    def test_missing_deal(self):
        text = render(classify("стадия 999999"), None, "")
        self.assertIn("999999", text)
        self.assertIn("Не нашёл", text)

    def test_calc_comment_is_plain(self):
        body = comment_body("[USER=1]Алексей[/USER]", "Посчитай лот", self.deal)
        self.assertIn("Алексей", body)
        self.assertNotIn("[USER", body)
        self.assertIn("1930", body)


if __name__ == "__main__":
    unittest.main()
