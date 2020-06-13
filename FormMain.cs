using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net.Mail;
using System.Net;
using System.Security.Permissions;

namespace Test
{
	[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
	[System.Runtime.InteropServices.ComVisibleAttribute(true)]
	public partial class FormMain : Form
	{
		struct Answer
		{
			public string text;
			public bool bCorrect;

			public Answer(string text, bool bCorrect)
			{
				this.text = text;
				this.bCorrect = bCorrect;
			}
		}

		struct Question
		{
			public string text;
			public int score;
			public bool bMulti;
			public List<Answer> answers;

			public Question(string text, int score, bool bMulti)
			{
				this.text = text;
				this.score = score;
				this.bMulti = bMulti;
				answers = new List<Answer>();
			}
		}

		List<Question> questions = new List<Question>();
		const string encryptionKey = "EncryptionKey";

		string mailSubject = "Тема теста не определена";
		string emailTo = "";
		string emailFrom = "";
		string emailPassword = "";
		string info = "";
		string startDateTime = "";
		string endDateTime = "";
		string userName = "";
		string scorePageContent = "";
		string mailBody = "";


		public FormMain()
		{
			InitializeComponent();
			webBrowser.ObjectForScripting = this;

			string[] lines = null;
			string content = "";
			string error = "";

			string fileNameEncrypt = Directory.GetCurrentDirectory() + "\\input.txt";
			string fileName = Directory.GetCurrentDirectory() + "\\data.bin";

			// Encrypt file if exist
			if (File.Exists(fileNameEncrypt))
			{
				string txt = Encryption.Encrypt(File.ReadAllText(fileNameEncrypt), encryptionKey);
				File.WriteAllText(fileName, txt);
			}

			// Load data from file
			if (File.Exists(fileName))
			{
				string txt = Encryption.Decrypt(File.ReadAllText(fileName), encryptionKey);
				lines = Regex.Split(txt, "\r\n|\r|\n");
			}
			else
			{
				error = "<div class='description'>Не удалось найти файл с тестом.</div>";
			}

			// Parse data
			foreach (var line in lines ?? Enumerable.Empty<string>())
			{
				char key = line.FirstOrDefault();
				switch (key)
				{
					// Header and description
					case '=':
						string header = line.Substring(1).TrimStart(' ');
						if (header.FirstOrDefault() == '=')
						{
							string description = header.Substring(1).TrimStart(' ');
							info += "<div class='description'>" + description + "</div>";
						}
						else
						{
							mailSubject = "[Результаты тестирования] " + header;
							info += "<div class='title'>" + header + "</div>";
						}
						break;

					// Email from and to
					case '@':
						string email = line.Substring(1).TrimStart(' ');
						if (email.FirstOrDefault() == '@')
						{
							emailFrom = email.Substring(1).TrimStart(' ');
						}
						else
						{
							emailTo = email;
						}
						break;

					// Email password
					case '#':
						emailPassword = line.Substring(1).TrimStart(' ');
						break;

					// Questions
					case '?':
						string question = line.Substring(1).TrimStart(' ');
						bool bMulti = question.FirstOrDefault() == '?';
						if (bMulti)
						{
							question = question.Substring(1).TrimStart(' ');
						}
						int score = 1;
						if (int.TryParse(question.FirstOrDefault().ToString(), out score))
						{
							question = question.Substring(1).TrimStart(' ');
						}
						questions.Add(new Question(question, score, bMulti));
						break;

					// Answers
					case '+':
					case '-':
						string answer = line.Substring(1).TrimStart(' ');
						questions.Last().answers.Add(new Answer(answer, key == '+'));
						break;
				}
			}

			// Build start page
			if (info != "")
			{
				content += "<div class='info'>" + info + "</div>";
			}

			if (content == "")
			{
				if (error == "")
				{
					error = "<div class='description'>Не удалось прочитать файл с тестом.</div>";
				}

				content +=
					"<div class='info'>" +
						"<div class='title'>Ошибка!</div>" + error +
						"<div class='description'>Чтобы зашифровать файл input.txt, поместите его в папку с программой.</div>" +
					"</div>";
			}
			else
			{
				content +=
					"<button class='button' name='start' onclick='window.external.StartTest()'>" +
						"<span>Начать тест</span>" +
					"</button>";
			}

			webBrowser.DocumentText = Properties.Resources.index.Replace("{0}", content);
		}

		public void StartTest()
		{
			startDateTime = DateTime.Now.ToString(new CultureInfo("ru-RU"));

			// Build questions page
			string content = "<div class='info'>" + info + "</div>";

			content +=
				"<form>" +
					"<div class='question'>ФИО, группа:</div>" +
					"<input type='text' autocomplete='off' id='nameInput' placeholder='Ваш ответ' />" +
				"</form>";

			for (int i = 0; i < questions.Count; i++)
			{
				var question = questions[i];
				content += "<form>";
				content += "<div class='question'>" + question.text + "</div>";
				for (int j = 0; j < question.answers.Count; j++)
				{
					var answer = question.answers[j];
					content += "<label class='checkcontainer' >" + answer.text;
					if (question.bMulti)
					{
						content += "<input type='checkbox' name='" + i + "' id='" + i + "_" + j + "'><span class='checkmark'></span>";
					}
					else
					{
						content += "<input type='radio' name='" + i + "' id='" + i + "_" + j + "'><span class='radiobtn'></span>";
					}
					content += "</label>";
				}
				content += "</form>";
			}

			content +=
				"<button class='button' name='confirm' onclick='window.external.EndTest()'>" +
					"<span>Завершить тест</span>" +
				"</button>";

			webBrowser.DocumentText = Properties.Resources.index.Replace("{0}", content);
		}

		public void EndTest()
		{
			// Check all questions answered
			foreach (HtmlElement form in webBrowser.Document.GetElementsByTagName("form"))
			{
				bool bValid = false;

				foreach (HtmlElement inputElement in form.GetElementsByTagName("input"))
				{
					if (inputElement.Id == "nameInput")
					{
						userName = inputElement.GetAttribute("value");
						if (!string.IsNullOrWhiteSpace(userName))
						{
							bValid = true;
							break;
						}
					}
					else if (inputElement.GetAttribute("checked") == "True")
					{
						bValid = true;
						break;
					}
				}

				form.Style = "";
				if (!bValid)
				{
					form.Style = "box-shadow: 0 0 0 2px #FF8A65, 0 0 24px rgba(0,0,0,.2);";
					webBrowser.Document.Window.ScrollTo(0, form.OffsetRectangle.Top - 32);
					return;
				}
			}

			endDateTime = DateTime.Now.ToString(new CultureInfo("ru-RU"));

			// Check answers and get result
			int score = 0;
			int maxScore = 0;
			for (int i = 0; i < questions.Count; i++)
			{
				var question = questions[i];
				bool bCorrect = false;

				if (question.bMulti)
				{
					bCorrect = true;
					for (int j = 0; j < question.answers.Count; j++)
					{
						var answer = question.answers[j];
						var check = webBrowser.Document.GetElementById(i + "_" + j);
						if (check != null)
						{
							bool bChecked = check.GetAttribute("checked") == "True";
							if (bChecked && !answer.bCorrect || !bChecked && answer.bCorrect)
							{

								bCorrect = false;
								break;
							}
						}
					}
				}
				else
				{
					for (int j = 0; j < question.answers.Count; j++)
					{
						var answer = question.answers[j];
						var check = webBrowser.Document.GetElementById(i + "_" + j);
						if (check != null)
						{
							bool bChecked = check.GetAttribute("checked") == "True";
							if (bChecked)
							{
								bCorrect = answer.bCorrect;
								break;
							}
						}
					}
				}

				maxScore += question.score;
				if (bCorrect)
				{
					score += question.score;
				}
			}

			int result = (int)(100.0 * score / maxScore);

			// Build score page
			string color = "#ff5722";
			if (result > 60)
				color = "#7cb342";
			else if (result > 40)
				color = "#2196f3";

			scorePageContent =
				"<div class='info'>Результаты отправлены на почту преподавателю</div>" +
				"<div class='title'>Ваш результат:</div>" +
				"<div class='description' style='color:" + color + ";'>" + score + " из " + maxScore + "</div>" +
				"<div class='description' style='color:" + color + ";'>" + result + "%</div>";

			// Build mail body
			mailBody =
				"<table>" +
					"<tr><td><b>ФИО:</b></td><td>" + userName + "</td></tr>" +
					"<tr><td><b>Время начала:</b></td><td>" + startDateTime + "</td></tr>" +
					"<tr><td><b>Время окончания:</b></td><td>" + endDateTime + "</td></tr>" +
					"<tr><td><b>Баллы:</b></td><td>" + score + " из " + maxScore + "</td></tr>" +
				"</table>";

			SendEmailAndShowScore();
		}

		public void SendEmailAndShowScore()
		{
			Cursor.Current = Cursors.WaitCursor;

			// Try to send email
			string error = "";
			try
			{
				MailMessage message = new MailMessage();
				message.From = new MailAddress(emailFrom);
				message.To.Add(new MailAddress(emailTo));
				message.Subject = mailSubject;
				message.Body = mailBody;
				message.IsBodyHtml = true;

				SmtpClient smtp = new SmtpClient();
				smtp.Port = 587;
				smtp.Host = "smtp.gmail.com";
				smtp.EnableSsl = true;
				smtp.UseDefaultCredentials = false;
				smtp.Credentials = new NetworkCredential(emailFrom, emailPassword);
				smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
				smtp.Send(message);
			}
			catch (Exception ex)
			{
				error = ex.Message;
			}

			// Show score page
			string content = scorePageContent;
			if (error != "")
			{
				content =
					"<div class='title'>Не удалось отправить результат</div>" +
					"<div class='info'>" + error + "</div>" +
					"<button class='button' name='resend' onclick='window.external.SendEmailAndShowScore()'><span>Отправить заново</span></button>";
			}

			webBrowser.DocumentText = Properties.Resources.score.Replace("{0}", content);

			Cursor.Current = Cursors.Default;
		}
	}
}
