using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using QlikView;
using System.Linq;
using System.Text.RegularExpressions;
using System.Drawing;

/*
	This file is part of Report Generator.

	Report Generator is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.

	Report Generator is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.

	You should have received a copy of the GNU General Public License
	along with Report Generator.  If not, see <http://www.gnu.org/licenses/>.
	Testin Git
 */

namespace GeneratorSpace
{
	public partial class GenerateReport : Form
	{
		//global variable dictionary for quick reference strings
		Dictionary<string, Tuple<string, string>> QuickRefVars = new Dictionary<string, Tuple<string, string>>();
		System.Diagnostics.Stopwatch stopWatch;

		public GenerateReport()
		{
			InitializeComponent();
		}

		private void GenerateReport_Load(object sender, EventArgs e)
		{
			//tooltip for helping users format their Word template
			//toolTip1.SetToolTip(this.btnHelp, "Standard Object Tags: <CH01>\nSelection Tags: <CH01{'FieldName','Selection'}>\nMultiple Selection Tags: <CH01{'FieldName1','Selection1','FieldName2','Selection2'}>\nLooping Tags: [FieldName]<CH01>[/FieldName]");

			//Initializing "global variable"
			ReportControl.staticSelections = new List<Tag>();
			ReportControl.listLog = this.lstLog;
		}

		private void GenerateReport_FormClosed(object sender, FormClosedEventArgs e)
		{
			//make sure all Word and Qlik documents and applications are closed
			exitWithGrace();

			//make sure the application closes when the form is closed
			System.Windows.Forms.Application.Exit();
		}

		/// <summary>
		/// makes selections, gets chart from QV, then pastes it to word.
		/// </summary>
		/// <param name="item">chart tag to fetch and paste</param>
		private bool GetChartsFromQV(string item, string quickRefLookup)
		{
			ReportControl.QVDoc.UnlockAll();
			ReportControl.QVDoc.ClearAll(true);
			Clipboard.Clear();
			string objectName = string.Empty;
			string fieldName = string.Empty;
			string selectionName = string.Empty;

			if (item.Contains("{"))//if the chart call contains a selection tag
			{
				int beginIndex = item.IndexOf('{');
				int endIndex = item.IndexOf('}');
				objectName = item.Substring(0, beginIndex);//everything between "<" and "{" is the chart name
				Console.WriteLine("Item: {0}", item);
				string controlString = item.Substring(beginIndex, item.Length - beginIndex);//this is the {field,selection,field2,selection...} format
				Console.WriteLine("Control string: {0}", controlString);
				List<Tag> selections = GeneratorSpace.Tag.interpretSelectionTag(controlString);
				bool validselections = applyQVSelections(selections);//apply selections before getting chart
				if (!validselections)//if the selections for a chart are not valid, don't paste the chart
				{
					Console.WriteLine("The selections were not valid for {0}", item);
					return false;
				}
			}
			else
			{
				objectName = item;

				applyQVSelections(new List<Tag>());//apply the static selections
			}

			SheetObject QVObject = ReportControl.QVDoc.GetSheetObject(objectName);//store QV object in memory to avoid costly QV queries

			if (QVObject != null)
			{
				Console.WriteLine("Object name: {0}. Field Name: {1}. Selection Name: {2}", objectName, fieldName, selectionName);

				switch (QVObject.GetObjectType())
				{
					case 11: //straight table

						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyTableToClipboard(true);
						Console.WriteLine("Found CH item: {0}", objectName);
						//pasteToWord(item);
						break;

					case 10: //pivot table
					  
						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyTableToClipboard(true);
						Console.WriteLine("Found CH item: {0}", objectName);
						
						//pasteToWord(item);
						break;

					case 1: //List Box

						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyTableToClipboard(true);
						Console.WriteLine("Found LB item: {0}", objectName);
						//pasteToWord(item);
						break;

					case 4: //Table Box
						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyTableToClipboard(true);
						Console.WriteLine("Found TB item: {0}", objectName);
						//pasteToWord(item);
						break;

					case 6: //text
						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyTextToClipboard();
						if (quickRefLookup != "")
						{
							QuickRefVars[quickRefLookup] = Tuple.Create<string, string>(item, Clipboard.GetText());
						}
						Console.WriteLine("Found TX item: {0}", objectName);
						//pasteToWord(item);
						break;

					default://everything else gets pasted as bitmap by default
						Console.WriteLine("ObjectType not Found for {0}, pasting as a bitmap", item);
						QVObject.GetSheet().Activate();
						ReportControl.QVApp.WaitForIdle();
						QVObject.CopyBitmapToClipboard();
						//pasteToWord(item);
						break;
				}
				return true;
			}
			else
			{
				Console.WriteLine("object {0} not found, moving on to the next one", objectName);
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("object " + objectName + " not found, moving on to the next one");
				return false;
			}
		}

		//take an area of text and a selection to add to each chart, return with selection added to each chart
		private string chartSelections(string text, string selection)
		{
			Regex rgx = new Regex("<[^>^<]*>");//match anything between "<>" brackets aside from other "<>" brackets
			Chart currChart;
			MatchCollection matchColle = rgx.Matches(text);
			Console.WriteLine(matchColle.Count);
			for (int i = 0; i < matchColle.Count; i++)//For each chart found in the text
			{
				Console.WriteLine(matchColle[i].Index);
				currChart = new Chart(matchColle[i].Value);
				currChart = currChart.AddSelectionTag(GeneratorSpace.Tag.interpretSelectionTag(selection));//make a new chart, with additional tags from function call
				text = text.Remove(matchColle[i].Index, matchColle[i].Length);//remove old text
				text = text.Insert(matchColle[i].Index, currChart.ToString());//add new text
				matchColle = rgx.Matches(text);//redo matches. possibly unnecessary after some changes.
			}
			return text;
		}

		//edited to open the QVdoc, also now edits the word document and deletes Looping tags in their place
		private bool EditWordForLooping(Tuple<string, string, int> loopStartTag, int loopEndIndex)
		{
			ReportControl.QVDoc.ClearAll();
			applyQVSelections(ReportControl.staticSelections);
			string loopField = loopStartTag.Item1; //QV field name
			string loopSelectionTag = loopStartTag.Item2; //selection tag including the brackets { . . . }
			int loopStartIndex = loopStartTag.Item3; //Word document index of beginning bracket [ of the loop tag

			Console.WriteLine("Field Name: {0}, Selection Tag: {1}, Start Index: {2}, End Index: {3}", loopField, loopSelectionTag, loopStartIndex, loopEndIndex);
			if (loopSelectionTag != "")
			{
				if (!applyQVSelections(GeneratorSpace.Tag.interpretSelectionTag(loopSelectionTag))) return false;
			}
			//store the text we want to copy
			string copyText;
			string pasteText;
			copyText = ReportControl.WordDoc.Range(loopStartIndex + loopField.Length + loopSelectionTag.Length + 2, loopEndIndex).Text;
			int deleteStart = loopStartIndex;
			int deleteEnd = loopEndIndex;
			Console.WriteLine("Delete beginning at {0}, delete ending at {1}", deleteStart, deleteEnd);
			Field bField = ReportControl.QVDoc.Fields(loopField);
			int copyEnd = loopStartIndex;
			if (bField != null)//making sure the field exists
			{
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Applying Looping selections.");
				Console.WriteLine("Applying Looping selections.");
				var selections = bField.GetPossibleValues();

				if (selections != null)//making sure the field has selections
				{
					lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of possible values: " + selections.Count);
					Console.WriteLine("Number of possible values: {0}", selections.Count);
					Clipboard.Clear();
					ReportControl.WordDoc.Range(deleteStart, deleteEnd).Delete();//remove old Looping tag and text

					for (int i = 0; i < selections.Count; i++)
					{
						lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Applying field selection to " + loopField + " with a value of " + selections[i].Text);
						Console.WriteLine("Applying field selection to {0} with a value of {1}", loopField, selections[i].Text);
						//pasteText = copyText.Trim().TrimEnd('>')+"{"+loopField+","+selections[i].Text+","+loopSelectionTag.Trim('{','}')+"}>\v";
						//add loop selection tags to the tag for the current iteration before passing it to the text editor
						Tag loopTag = new GeneratorSpace.Tag(loopField, selections[i].Text);
						List<Tag> allTags = GeneratorSpace.Tag.interpretSelectionTag(loopSelectionTag);
						allTags = GeneratorSpace.Tag.addTag(allTags, loopTag);
						pasteText = chartSelections(copyText, GeneratorSpace.Tag.listTagsToString(allTags));

						//new chart tag complete, now paste it to word
						Clipboard.SetText(pasteText);
						ReportControl.WordDoc.Range(copyEnd, copyEnd).Paste();
						copyEnd += pasteText.Length;
					}
					ReportControl.WordDoc.Range(copyEnd, copyEnd + loopField.Length + 5).Delete();//remove the end of the Looping tag
					Clipboard.Clear();
					return true;
				}//field has no possible selections
				return false;
			}
			else//field does not exist
			{
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Field name not found: " + loopField);
				Console.WriteLine("Field name not found: {0}", loopField);
				return false;
			}
		}

		//checks syntax of interpretSelectionTag and turns a raw string of format {field1,selection1,field2,selection2...} into touples
		//DEPRICATED, use Tag.interpretSelectionTag instead
		public static List<Tuple<string, string>> interpretSelectionTag(string selectionTag)
		{
			Console.WriteLine("Selection tag to interpret: {0}", selectionTag);
			List<Tuple<string, string>> selections = new List<Tuple<string, string>>();
			if (selectionTag != "")
			{
				if (Equals(selectionTag[0].ToString(), "{") && Equals(selectionTag[selectionTag.Length - 1].ToString(), "}"))
				{
					int commaCount = selectionTag.Count(i => i == ',');

					if (commaCount % 2 != 0) //properly formatted tags will always have an odd number of commas
					{

						//error checking complete, now manipulate strings to make list
						while (selectionTag.Contains(","))
						{
							int commaIndex;
							string fieldName, fieldValue;

							commaIndex = selectionTag.IndexOf(","); //get index of first comma
							fieldName = selectionTag.Substring(1, commaIndex - 1);

							if (selectionTag.Count(i => i == ',') == 1)
							{
								selectionTag = selectionTag.Remove(1, commaIndex);
								fieldValue = selectionTag.Substring(1, selectionTag.Length - 2);
								selectionTag = selectionTag.Remove(1, fieldValue.Length);
							}
							else
							{
								selectionTag = selectionTag.Remove(1, fieldName.Length + 1);
								commaIndex = selectionTag.IndexOf(",");
								fieldValue = selectionTag.Substring(1, commaIndex - 1);
								selectionTag = selectionTag.Remove(1, fieldValue.Length + 1);
							}

							selections.Add(Tuple.Create(fieldName.Trim(), fieldValue.Trim()));
						}

						return selections;
					}
					else
					{
						Console.WriteLine("Selection tag invalid, incorrect number of arguments: {0}", selectionTag);
						return null;
					}
				}
				else
				{
					Console.WriteLine("Selection tag invalid, improper brace structure: {0}", selectionTag);
					return null;
				}
			}
			else
			{
				Console.WriteLine("Selection tag was null");
				return selections;
			}
		}

		//removed rollback list, added selectvalues functionality for '|' operator, works for numerics now but very slow
		/// <summary>
		/// Apply selections in the QlikView document by looping through a list of selections
		/// </summary>
		/// <param name="selectionList">List if Tuples containing fieldName and fieldValue data</param>
		private bool applyQVSelections(List<Tag> selectionList)
		{

			if (selectionList != null) //make sure there is data in the list
			{
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Applying QV selections.");
				Console.WriteLine("Applying QV selections.");
				//aggregate static selections and selections to apply into a single list
				List<Tag> selectionTagList = new List<Tag>();
				foreach (var sel in selectionList)
				{
					selectionTagList.Add(sel);
				}
				if (ReportControl.staticSelections != null)
				{
					foreach (var statTag in ReportControl.staticSelections)
					{
						selectionTagList.Add(statTag);
					}
				}
				selectionTagList = GeneratorSpace.Tag.aggTags(selectionTagList);//all selections to be made for the current chart
				Console.WriteLine(GeneratorSpace.Tag.listTagsToString(selectionTagList));

				foreach (var item in selectionTagList) //loop through items
				{
					if (ReportControl.QVDoc.Fields(item.Field) != null) //if field exists...
					{
						string fieldValue = item.Selection;
						int argCount = 1; //counter for number of arguments in fieldValue

						//QVDoc.Evalaute wants multiple selections in a csv format rather than normal selection syntax
						if (fieldValue.Contains("|"))
						{
							argCount = fieldValue.Count(i => i == '|') + 1;
							fieldValue = fieldValue.Replace("|", "','");
						}

						//build the QlikView formula ahead of time
						string countString = @"COUNT(DISTINCT{$<[" + item.Field + @"]={'" + fieldValue + @"'}>}[" + item.Field + @"])";
						fieldValue = fieldValue.Replace("','", "|");//put the line back in instead of the comma

						//parse the formula result as an integer and make sure it's equal to the argCount
						//this verifies that ALL of the fieldValues exist in the fieldName
						if (int.Parse(ReportControl.QVDoc.Evaluate(countString)) == argCount)
						{
							if (fieldValue.Contains("|"))
							{//if there are multiple select targets on one field we have to perform them at the same time

								string selection = item.Selection;
								string[] selList = selection.Split('|');
								IArrayOfFieldValue multiSelect = ReportControl.QVDoc.Fields(item.Field).GetNoValues();
								Field mField = ReportControl.QVDoc.Fields(item.Field);
								for (int i = 0; i < selList.Length; i++)
								{
									lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Applying selection to field " + item.Field + " with value " + selList[i]);
									Console.WriteLine("Applying selection to field {0} with value {1}", item.Field, selList[i]);
									multiSelect.Add();
									multiSelect[i].Text = selList[i].Trim();
									if (mField.GetProperties().IsNumeric)//if it is numeric, we have to set the number
									{
										mField.Select(selList[i].Trim());
										multiSelect[i].IsNumeric = true;
										multiSelect[i].Number = mField.GetSelectedValues()[0].Number;//QV does not have a "get value" function afaik, this is slow but works
									}
								}
								ReportControl.QVDoc.Fields(item.Field).SelectValues(multiSelect);
							}
							else
							{
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Applying selection to field " + item.Field + " with value " + item.Selection);
								Console.WriteLine("Applying selection to field {0} with value {1}", item.Field, item.Selection);
								ReportControl.QVDoc.Fields(item.Field).Select(item.Selection);
							}
						}
						else
						{
							lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Selection value not found: " + item.Selection);
							Console.WriteLine("Selection value not found: {0}", item.Selection);
							return false;
						}
					}
					else
					{
						lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Selection field name not found: " + item.Field);
						Console.WriteLine("Selection field name not found: {0}", item.Field);
						return false;
					}
				}
			}
			return true;
		}

		//given a chart tag (eg. "CH01") pastes it to word by parsing for the chart
		//TODO: rework this to not have to parse the document to paste
		private void pasteToWord(Word.Range wordSelection, Dictionary<string,string> parameters = null)
		{
			if (Clipboard.ContainsData(DataFormats.Bitmap))//check data format and paste
			{
				lstLog.TopIndex = lstLog.Items.Count - 1;
				lstLog.Items.Add("Chart In Clipboard");
				Console.WriteLine("Chart Found In Clipboard");
				if (parameters != null)
				{
					Bitmap oldImage = (Bitmap)Clipboard.GetData(DataFormats.Bitmap);
					double ratio = (double)oldImage.Height / oldImage.Width;
					if (ratio == 0.0) ratio = 1.0;
					int oldWidth = oldImage.Width;
					int oldHeight = oldImage.Height;

					double tnewWidth = oldWidth;
					double tnewHeight = oldHeight;

					int newWidth = oldWidth;
					int newHeight = oldHeight;


					if (parameters.ContainsKey("HEIGHT") && parameters.ContainsKey("WIDTH"))
					{
						tnewWidth = Convert.ToDouble(parameters["WIDTH"]);
						newWidth = (int)(oldImage.HorizontalResolution * tnewWidth);
						tnewHeight = Convert.ToDouble(parameters["HEIGHT"]);
						newHeight = (int)(oldImage.VerticalResolution * tnewHeight);

					}
					else if (parameters.ContainsKey("WIDTH"))
					{
						tnewWidth = Convert.ToDouble(parameters["WIDTH"]);
						newWidth = (int)(oldImage.HorizontalResolution * tnewWidth);
						newHeight = (int)(newWidth * ratio);
					}
					else if (parameters.ContainsKey("HEIGHT"))
					{
						tnewHeight = Convert.ToDouble(parameters["HEIGHT"]);
						newHeight = (int)(oldImage.VerticalResolution * tnewHeight);
						newWidth = (int)(newHeight / ratio);
					}
					else
					{
						wordSelection.Paste();
						return;
					}

					Bitmap newImage = new Bitmap(oldImage, new Size(newWidth, newHeight));
					Clipboard.SetData(DataFormats.Bitmap, newImage);
					//wordSelection.PasteSpecial(0, false, Word.WdOLEPlacement.wdInLine, false, Word.WdPasteDataType.wdPasteEnhancedMetafile);
					wordSelection.Paste();
				}
				else
				{
					wordSelection.PasteSpecial(0, false, Word.WdOLEPlacement.wdInLine, false, Word.WdPasteDataType.wdPasteEnhancedMetafile);
				}
			}
			else if (Clipboard.ContainsText())
			{
				lstLog.TopIndex = lstLog.Items.Count - 1;
		lstLog.Items.Add("Text Object Found In Clipboard");
				
				Console.WriteLine("Text Object Found In Clipboard");

				var htm = Clipboard.GetData(DataFormats.Html);
				if (htm != null && htm.ToString().Contains("<META CONTENT=\"PivotTable\">") && htm.ToString().Contains("*"))
				{
					string argue = formatPivotTable(htm.ToString());
			string argue1 = argue.ToString().Replace("<TABLE ", "<TABLE align=\"center\" ");
					String mm = argue1.Replace("style=\"", "style=\"border-collapse:collapse; ");
					CopyToClipboard(mm);
					//Clipboard.SetData(DataFormats.Html, argue); //this may work sometimes but its not reliable
				}
				else if (htm != null && htm.ToString().Contains("<META CONTENT=\"PivotTable\">"))
				{
					string argue = htm.ToString().Replace("<TABLE ", "<TABLE align=\"center\" ");
			string argue1 = argue.ToString().Replace("<TABLE ", "<TABLE align=\"center\" ");
					String mm = argue1.Replace("style=\"", "style=\"border-collapse:collapse; ");
					CopyToClipboard(mm);
					//Clipboard.SetData(DataFormats.Html, argue); //this may work sometimes but its not reliable
				} 
		else if (htm != null && htm.ToString().Contains("<TABLE"))
				{
					string argue = htm.ToString().Replace("<TABLE ", "<TABLE align=\"center\" ");
					String mm = argue.Replace("style=\"", "style=\"border-collapse:collapse; ");
					CopyToClipboard(mm);
					//Clipboard.SetData(DataFormats.Html, argue); //this may work sometimes but its not reliable

					wordSelection.Paste();
				}
				else
				{
					wordSelection.Paste();
				}
			}
		}

		private string formatPivotTable(string htm)
		{
			String neww3 = htm.ToString().Replace("<TD BGCOLOR=\"#f5f5f5\">&nbsp\t", "");
			String neww1 = neww3.ToString().Replace("<TD BGCOLOR=\"#ffffff\">&nbsp\t", "");
			String neww2 = neww1.ToString().Replace("<TH BGCOLOR=\"#f5f5f5\"><FONT COLOR=\"#363636\"><B>*<B></B></FONT>\t", "");
			String mm = neww2.Replace("style=\"", "style=\"border-collapse:collapse; ");
			String neww4 = mm.Replace("<TABLE ", "<TABLE align=\"center\" ");
			String[] spli = neww4.Split(new string[] { "<TR " }, StringSplitOptions.None);
			if (spli.Length > 1)
			{
				string proc = spli[1];
				int tds = proc.Split(new string[] { "<TD " }, StringSplitOptions.None).Length - 1;
				double perTD;
				if (tds > 0)
				{
					perTD = 75.0 / tds;
				}
				else
				{
					perTD = 75.0;
				}
				string ss = $"<TD width=\"{perTD}%\" ";
				string nw = proc.Replace("<TD ", ss);
				spli[1] = nw;
				string ret = String.Join("<TR ", spli);
				return ret;
			}
			return neww4;
		}

		private void CopyToClipboard(string fullHtmlContent)
		{
			// http://pavzav.blogspot.com/2010/11/how-to-copy-html-content-to-clipboard-c.html
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			string header = @"Version:1.0
							StartHTML:<<<<<<<1
							EndHTML:<<<<<<<2
							StartFragment:<<<<<<<3
							EndFragment:<<<<<<<4";
			sb.Append(header);
			int startHTML = sb.Length;
			sb.Append(fullHtmlContent);
			int endHTML = sb.Length;
   
			sb.Replace("<<<<<<<1", To8CharsString(startHTML));
			sb.Replace("<<<<<<<2", To8CharsString(endHTML));
			sb.Replace("<<<<<<<3", To8CharsString(startHTML));
			sb.Replace("<<<<<<<4", To8CharsString(endHTML));
   
			Clipboard.Clear();
			Clipboard.SetText(sb.ToString(), TextDataFormat.Html);
		}
		private static string To8CharsString(int x)
		{
			  return x.ToString("0#######");
		}

	//**NEW**parse file and process triangle bracketed tags
		private Tuple<int, int> findCharts(Word.Range content)
		{
			Console.WriteLine("Range search initiatied.");

			bool temp;
			int successes = 0;
			int failures = 0;
			var wordSelection = content;
			wordSelection.Find.ClearFormatting();
			wordSelection.Find.Text = "[<][!<>]{1,}[>]"; //find anything in between brackets
			wordSelection.Find.Forward = true;
			wordSelection.Find.Wrap = Word.WdFindWrap.wdFindStop;
			wordSelection.Find.Format = false;
			wordSelection.Find.MatchCase = false;
			wordSelection.Find.MatchWholeWord = false;
			wordSelection.Find.MatchWildcards = true;
			wordSelection.Find.MatchSoundsLike = false;
			wordSelection.Find.MatchAllWordForms = false;
			
			List<string> found = new List<string>();
			string str;
			string tagText;
			char[] trimchar = new char[2];
			trimchar[0] = '<';
			trimchar[1] = '>';
			while (wordSelection.Find.Execute())
			{
				if (wordSelection.Find.Found)
				{
					str = wordSelection.Text;
					tagText = str.Trim(trimchar);
					List<Tuple<string, string>> QuickRefValue = new List<Tuple<string, string>>();
					bool passOnThru = false;
					string dictLookup = string.Empty;

					if (tagText[0].ToString() == "!")//first character is ! denoting quickrefvar
					{
						Console.WriteLine("Found quick reference tag: {0}", tagText);
						Clipboard.Clear();
						QuickRefValue = QuickRefVars.Where(x => x.Key.Equals(tagText.Substring(1))).Select(x => x.Value).ToList();

						if (!(QuickRefValue.Count == 0))
						{
							if (QuickRefValue[0].Item2 != "")
							{
								Clipboard.SetText(QuickRefValue[0].Item2);
								Console.WriteLine("Applying from quick reference: {0}, {1}", tagText, QuickRefValue[0].Item2);
								wordSelection.Paste();
								passOnThru = true;
							}
							else
							{
								dictLookup = tagText.Substring(1);
								tagText = QuickRefValue[0].Item1;
							}
						}
					}

					if (!passOnThru)
					{
						string paramStr = "";
						string tagStr = "";
						string[] paramExists = new string[0];

						if (tagText.Contains("?"))
						{
							paramExists = tagText.Split('?');
							tagStr = paramExists[0];
							paramStr = paramExists[1];
						}
						else
						{
							tagStr = tagText;
						}

						temp = GetChartsFromQV(tagStr, dictLookup);

						if (temp && paramExists.Length > 1)
						{
							List<string> paramArr = new List<string>();

							if (paramStr.Contains('&'))
							{
								paramArr = paramStr.Split('&').ToList();
							}
							else
							{
								paramArr.Add(paramStr);
							}

							Dictionary<string, string> attributes = new Dictionary<string, string>();

							foreach (string par in paramArr)
							{
								string[] index = new string[2];
								string parAtt = "";
								string parVal = "";
								if (par.Contains('='))
								{
									index = par.Split('=');
									parAtt = index[0];
									parVal = index[1];
								}

								switch (parAtt.ToUpper())
								{
									case "HEIGHT":
										try
										{
											Convert.ToDouble(parVal);
											attributes.Add("HEIGHT", parVal);
										}
										catch
										{
											break;
										}
										break;
									case "WIDTH":
										try
										{
											Convert.ToDouble(parVal);
											attributes.Add("WIDTH", parVal);
										}
										catch
										{
											break;
										}
										break;
									default:
										break;
								}
							}
							pasteToWord(wordSelection, attributes);
							successes++;
						}
						else if (temp)
						{
							pasteToWord(wordSelection);
							successes++;
						}
						else
						{
							failures++;
						}
					}

					wordSelection.Start = wordSelection.End;
				}
			}

			replaceQuickReferenceTagsInHeadersAndFooters(wordSelection);

			return new Tuple<int, int>(successes, failures);
		}

		private void replaceQuickReferenceTagsInHeadersAndFooters(Word.Range wordSelection)
		{
			// helpful link: http://stackoverflow.com/questions/17714642/replace-field-in-headerfooter-in-word-using-interop
			foreach (Word.Section section in ReportControl.WordDoc.Sections)
			{
				//ReportControl.WordDoc.TrackRevisions = false;//Disable Tracking for the Field replacement operation
				//Get all Headers
				Word.HeadersFooters headers = section.Headers;
				Word.HeadersFooters footers = section.Headers;

				replaceQuickReferenceTagsInHeadersAndFooters(section.Headers);
				replaceQuickReferenceTagsInHeadersAndFooters(section.Footers);
			}
		}

		private void replaceQuickReferenceTagsInHeadersAndFooters(Word.HeadersFooters headersOrFooters)
		{
			//Section headerfooter loop for all types enum WdHeaderFooterIndex. wdHeaderFooterEvenPages/wdHeaderFooterFirstPage/wdHeaderFooterPrimary;                          
			foreach (Microsoft.Office.Interop.Word.HeaderFooter header in headersOrFooters)
			{
				string headerText = header.Range.Text;
				var myStoryRange = header.Range;
				string tagText;
				if (headerText.Contains('<') && headerText.Contains('>') && headerText.Length > 3)
				{
					tagText = headerText.Split('<', '>')[1];
				}
				else
				{
					tagText = string.Empty;
				}

				List<Tuple<string, string>> QuickRefValue = new List<Tuple<string, string>>();
				if (tagText != string.Empty && tagText[0].ToString() == "!")
				{
					Console.WriteLine("Found quick reference tag: {0}", tagText);
					Clipboard.Clear();
					QuickRefValue = QuickRefVars.Where(x => x.Key.Equals(tagText.Substring(1))).Select(x => x.Value).ToList();

					if (!(QuickRefValue.Count == 0))
					{
						if (QuickRefValue[0].Item2 != "")
						{
							//Clipboard.SetText(QuickRefValue[0].Item2);
							Console.WriteLine("Applying from quick reference: {0}, {1}", tagText, QuickRefValue[0].Item2);
							FindAndReplace(ReportControl.WordDoc, "<" + tagText + ">", QuickRefValue[0].Item2);
							//wordSelection.Paste();
						}
					}
					//wordSelection.Start = wordSelection.End;
				}
				else
				{
					Console.WriteLine("There is something other than a quick reference tag in a header");
				}
			}
		}


		private void FindAndReplace(Word.Document document, string placeHolder, string newText)
		{
			// this was taken from here: http://forum.katarincic.com/default.aspx?g=posts&m=456
			object missingObject = null;
			object item = Word.WdGoToItem.wdGoToPage;

			object whichItem = Word.WdGoToDirection.wdGoToFirst;
			object replaceAll = Word.WdReplace.wdReplaceAll;
			object forward = true;
			object matchAllWord = true;
			object matchCase = false;
			object originalText = placeHolder;
			object replaceText = newText;

			document.GoTo(ref item, ref whichItem, ref missingObject, ref missingObject);
			foreach (Word.Range rng in document.StoryRanges)
			{
				rng.Find.Execute(ref originalText, ref matchCase,
				ref matchAllWord, ref missingObject, ref missingObject, ref missingObject, ref forward,
				ref missingObject, ref missingObject, ref replaceText, ref replaceAll, ref missingObject,
				ref missingObject, ref missingObject, ref missingObject);
			}
		}

		private void preProcessing(Word.Range content)
		{
			var wordSelection = content;
			wordSelection.Find.ClearFormatting();
			wordSelection.Find.Text = "[<][=]*[>]"; //find anything in between brackets
			wordSelection.Find.Forward = true;
			wordSelection.Find.Wrap = Word.WdFindWrap.wdFindStop;
			wordSelection.Find.Format = false;
			wordSelection.Find.MatchCase = false;
			wordSelection.Find.MatchWholeWord = false;
			wordSelection.Find.MatchWildcards = true;
			wordSelection.Find.MatchSoundsLike = false;
			wordSelection.Find.MatchAllWordForms = false;

			string str; //found text
			while (wordSelection.Find.Execute())
			{
				if (wordSelection.Find.Found)
				{
					str = wordSelection.Text;
					//Console.WriteLine(str);
					//bracketed statement found, add to list to get charts
					str = str.Replace("[", "#%11111");
					str = str.Replace("]", "#%22222");
					str = str.Replace('“', '"');
					str = str.Replace('”', '"');
					str = str.Replace('‘', '\'');
					str = str.Replace('’', '\'');
					wordSelection.Text = str;
				}
			}
		}

		private void postProcessing(Word.Range content)
		{
			var wordSelection = content;
			wordSelection.Find.ClearFormatting();
			wordSelection.Find.Text = "[<][=]*[>]"; //find anything in between brackets
			wordSelection.Find.Forward = true;
			wordSelection.Find.Wrap = Word.WdFindWrap.wdFindStop;
			wordSelection.Find.Format = false;
			wordSelection.Find.MatchCase = false;
			wordSelection.Find.MatchWholeWord = false;
			wordSelection.Find.MatchWildcards = true;
			wordSelection.Find.MatchSoundsLike = false;
			wordSelection.Find.MatchAllWordForms = false;

			string str; //found text
			while (wordSelection.Find.Execute())
			{
				if (wordSelection.Find.Found)
				{
					str = wordSelection.Text;
					//Console.WriteLine(str);
					//bracketed statement found, add to list to get charts
					str = str.Replace("#%11111", "[");
					str = str.Replace("#%22222", "]");
					wordSelection.Text = str;
				}
			}
		}

		//shape overload for findCharts
		private Tuple<int, int> findChartsInShapes(Word.Shape shape)
		{
			Console.WriteLine("Shape search initiated.");

			if (shape.TextFrame.HasText >= 0) return new Tuple<int, int>(0, 0);
			bool temp;
			int successes = 0;
			int failures = 0;

			var wordSelection = shape.TextFrame.ContainingRange;
			wordSelection.Find.ClearFormatting();
			wordSelection.Find.Text = "[<][!<>]{1,}[>]"; //find anything in between brackets
			wordSelection.Find.Forward = true;
			wordSelection.Find.Wrap = Word.WdFindWrap.wdFindStop;
			wordSelection.Find.Format = false;
			wordSelection.Find.MatchCase = false;
			wordSelection.Find.MatchWholeWord = false;
			wordSelection.Find.MatchWildcards = true;
			wordSelection.Find.MatchSoundsLike = false;
			wordSelection.Find.MatchAllWordForms = false;

			List<string> found = new List<string>();
			string str;
			string tagText;
			char[] trimchar = new char[2];
			trimchar[0] = '<';
			trimchar[1] = '>';

			while (wordSelection.Find.Execute())
			{
				if (wordSelection.Find.Found)
				{
					str = wordSelection.Text;
					tagText = str.Trim(trimchar);
					List<Tuple<string, string>> QuickRefValue = new List<Tuple<string, string>>();
					bool passOnThru = false;
					string dictLookup = string.Empty;

					if (tagText[0].ToString() == "!")//first character is ! denoting quickrefvar
					{
						Console.WriteLine("Found quick reference tag: {0}", tagText);
						Clipboard.Clear();
						QuickRefValue = QuickRefVars.Where(x => x.Key.Equals(tagText.Substring(1))).Select(x => x.Value).ToList();

						if (!(QuickRefValue.Count == 0))
						{
							if (QuickRefValue[0].Item2 != "")
							{
								Clipboard.SetText(QuickRefValue[0].Item2);
								Console.WriteLine("Applying from quick reference: {0}, {1}", tagText, QuickRefValue[0].Item2);
								wordSelection.Paste();
								passOnThru = true;

							}
							else
							{
								dictLookup = tagText.Substring(1);
								tagText = QuickRefValue[0].Item1;
							}
						}
					}

					if (!passOnThru)
					{
						temp = GetChartsFromQV(tagText, dictLookup);

						if (temp)
						{
							pasteToWord(wordSelection);
							successes++;
						}
						else
						{
							failures++;
						}
					}

					wordSelection.Start = wordSelection.End;
				}
			}
			return new Tuple<int, int>(successes, failures);
		}

		/// <summary>
		/// Run through the Word document and find all Looping tags to interpret as selection tags
		/// </summary>
		private Tuple<int, int, int, int> findLooping()
		{
			//open the Word document (which is surely closed as verified by the btnGenerate_Click method)
			int success = 0;
			int failed = 0;
			int falseStarts = 0;
			int falseEnds = 0;
			try
			{
				var wordSelection = ReportControl.WordDoc.Content;
				wordSelection.Find.ClearFormatting();
				wordSelection.Find.Text = "[[]*[]]"; //find anything in between brackets
				wordSelection.Find.Forward = true;
				wordSelection.Find.Wrap = Word.WdFindWrap.wdFindStop;
				wordSelection.Find.Format = false;
				wordSelection.Find.MatchCase = false;
				wordSelection.Find.MatchWholeWord = false;
				wordSelection.Find.MatchWildcards = true;
				wordSelection.Find.MatchSoundsLike = false;
				wordSelection.Find.MatchAllWordForms = false;

				//initialize local variables
				List<Tuple<string, string, int>> startList = new List<Tuple<string, string, int>>();
				string strWord;

				while (wordSelection.Find.Execute()) //while running through the word document...
				{
					if (wordSelection.Find.Found) //if something is found which is enclosed by brackets...
					{
						//temporarily store the text that was found, includes the brackets
						strWord = wordSelection.Text;

						//if the second character is a "/" then this is the ending piece of the Looping tag
						if (strWord.Substring(0, 2).Contains("/"))
						{
							//extract the field name and discard the brackets and "/"
							string loopField = strWord.Substring(2, strWord.Length - 3);

							if (startList != null) //make sure there is content in the startlist
							{
								//if the field is has a corresponding loop tag in the startlist...
								if (startList.Exists(i => i.Item1.Equals(loopField)))
								{
									//get the index in the list of the corresponding record
									int index = startList.FindIndex(i => i.Item1.Equals(loopField));

									lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("loop end found: " + loopField);
									Console.WriteLine("loop end found: {0}", loopField);

									//pass the tuple to the EditWordForLooping method along with the Word index of the
									// "[" in the ending loop tag
									bool sucFlag = EditWordForLooping(startList[index], wordSelection.Start);//returns true for success, false for fail
									if (sucFlag)
									{
										wordSelection.End = wordSelection.Start;//reset the selected text
										success++;
									}
									else failed++;

									//remove the record from the list as not to confuse the remainder of the scan
									startList.RemoveAt(index);
								}
								else
								{
									lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Looping end found: " + loopField + ", but start not found in the list. Moving on.");
									Console.WriteLine("Looping end found: {0}, but start not found in the list. Moving on.", loopField);
									falseEnds++;
								}
							}
							else
							{
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Looping end found: " + loopField + ", but no starts detected. Moving on.");
								Console.WriteLine("Looping end found: {0}, but no starts detected. Moving on.", loopField);
								falseEnds++;
							}
						}
						else //no "/" is found, this is the start of the Looping tag
						{
							string selectionTag = "";

							//check if the Looping tag contains a selection tag
							if (strWord.Contains("{"))
							{
								int index = strWord.IndexOf("{");

								//extract the selection tag
								selectionTag = strWord.Substring(index, strWord.Length - index - 1);

								//extract the fieldname
								strWord = strWord.Substring(1, strWord.Length - selectionTag.Length - 2);

								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Selection tag found");
								Console.WriteLine("Selection tag found");
							}
							else
							{
								//drop the brackets so it is just the field name
								strWord = strWord.Substring(1, strWord.Length - 2);
							}

							//load start list with the fieldname, selectionTag (which can be ""), and the index of the "[" that 
							//started the Looping tag
							startList.Add(Tuple.Create(strWord, selectionTag, wordSelection.Start));

							lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("loop start found: " + strWord + ", Selection tag: " + selectionTag);
							Console.WriteLine("loop start found: {0}, Selection tag: {1}", strWord, selectionTag);
						}
					}
				}
				falseStarts += startList.Count();
			}
			catch (Exception ex) //Word document was inaccessible to the scan for some reason
			{
				MessageBox.Show("Something went wrong: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add(ex.StackTrace.ToString());
				Console.WriteLine(ex.StackTrace.ToString());
			}
			finally
			{
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Preparing to exit.");
				Console.WriteLine("Preparing to exit.");

				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("############## Looping Results ###################");
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of successfully processed Looping tags: " + success.ToString());
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of unsuccessfully processed Looping tags: " + failed.ToString());
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of start tags with no end: " + falseStarts.ToString());
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of end tags with no start: " + falseEnds.ToString());
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("##############################################");
			}
			return new Tuple<int, int, int, int>(success, failed, falseStarts, falseEnds);
		}

		/// <summary>
		/// Inititates the process of creating the report
		/// </summary>
		private void btnGenerate_Click(object sender, EventArgs e)
		{
			//initialize local variables and clear the on-screen log
			string wordPath = txtWordPath.Text;
			string qlikPath = txtQlikPath.Text;
			lstLog.Items.Clear();
			
			stopWatch = new System.Diagnostics.Stopwatch();
			stopWatch.Start();

			if (wordPath != "" && qlikPath != "") //process won't work unless both paths are specified
			{
				if (validatePaths())
				{
					if (isFileOpen(wordPath)) //program will fail if Word doc is already open, the state of the QlikView doc doesn't matter
					{
						MessageBox.Show("Word document is open. Please close before continuing.", "File Open", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
					else
					{
						try
						{
							//load global variables with file paths that have passed verification
							ReportControl.wordPath = wordPath;
							ReportControl.qlikPath = qlikPath;
							openWordDocument();
							openQlikDocument();

							//validate any selection tags in the static selection tag text box
							if (applyQVSelections(GeneratorSpace.Tag.interpretSelectionTag(txtStaticSelections.Text.Trim())) && GeneratorSpace.Tag.interpretSelectionTag(txtStaticSelections.Text.Trim()) != null)
							{
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Static selection tag validated. Beginning Word search process.");
								Console.WriteLine("Static selection tag validated. Beginning Word search process.");
								ReportControl.staticSelections = GeneratorSpace.Tag.interpretSelectionTag(txtStaticSelections.Text.Trim());//save static selections

								preProcessing(ReportControl.WordDoc.Content);       //removes square brackets from charts so no false positives during Looping
								Tuple<int, int, int, int> BResults;
								BResults = findLooping();
								postProcessing(ReportControl.WordDoc.Content);      //puts back square brackets after Looping

								Tuple<int, int> CResults;
								int boxsuccess = 0;
								int boxfail = 0;
								CResults = findCharts(ReportControl.WordDoc.Content);
								boxsuccess += CResults.Item1;
								boxfail += CResults.Item2;
								Console.WriteLine(ReportControl.WordDoc.Shapes.Count);

								//lastly, we need to check all of the text boxes and other shapes in the word document for any tags
								foreach (Word.Shape shape in ReportControl.WordDoc.Shapes)
								{
									CResults = findChartsInShapes(shape);
									boxsuccess += CResults.Item1;
									boxfail += CResults.Item2;
								}

								CResults = new Tuple<int, int>(boxsuccess, boxfail);

								stopWatch.Stop();
								TimeSpan runTime = stopWatch.Elapsed;

								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("############## Looping Results ##############");
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of successfully processed Looping tags: " + BResults.Item1.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of unsuccessfully processed Looping tags: " + BResults.Item2.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of start tags with no end: " + BResults.Item3.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of end tags with no start: " + BResults.Item4.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("##############################################");

								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("############## Chart Retrieval Breakdown ##############");
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of Charts successfully pasted to word: " + CResults.Item1.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Number of Invalid Chart tags: " + CResults.Item2.ToString());
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Run Time: " + runTime.TotalMinutes + " minutes");
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("############## END ##############");
								lstLog.SelectedIndex = lstLog.Items.Count - 1;

								exitWithGrace();
							}
							else
							{
								lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Static selection tag could not be validated. Cannot proceed until resolved.");
								Console.WriteLine("Static selection tag could not be validated. Cannot proceed until resolved.");

								//the if statement opens the qlik document, this ensures that it is closed in this path
								exitWithGrace();
							}
						} catch (Exception ee)
						{
							lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add($"A(n) {ee.GetType().Name} has caused the program to close");
							Console.WriteLine($"A(n) {ee.GetType().Name} has caused the program to close");
							exitWithGrace();
						}
					}
				}
			}
			else
			{
				MessageBox.Show("Use the Browse buttons to find the Word and Qlik files", "Missing Paths", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		/// <summary>
		/// Returns boolean variable based on whether the file is open
		/// </summary>
		/// <param name="filePath">File path of the Word document of which we want to check the status</param>
		private bool isFileOpen(string filePath)
		{
			//get info about the file
			FileInfo file = new FileInfo(filePath);
			FileStream stream = null;

			try
			{
				//if the file is open, this will fail and move to the catch block
				stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
			}
			catch (IOException)
			{
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("File at path " + filePath + " confirmed to be open.");
				Console.WriteLine("File at path {0} confirmed to be open.", filePath);
				return true;
			}
			finally
			{
				//clean up if something screwy happened
				if (stream != null) stream.Close();
			}

			lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("File at path " + filePath + " confirmed to be closed");
			Console.WriteLine("File at path {0} confirmed to be closed", filePath);
			return false;
		}

		/// <summary>
		/// Opens the Qlik application and document associated with the global variable "qlikPath"
		/// </summary>
		private void openQlikDocument()
		{
			//need the QlikView reference in order for this to work properly
			Type QlikViewApp = Type.GetTypeFromProgID("QlikTech.QlikView");
			ReportControl.QVApp = Activator.CreateInstance(QlikViewApp) as QlikView.Application;
			ReportControl.QVDoc = ReportControl.QVApp.OpenDoc(ReportControl.qlikPath);

			//unlock all fields in the document and clear all selections
			ReportControl.QVDoc.UnlockAll();
			ReportControl.QVDoc.ClearAll(true);

			if (cbxQlikReload.Checked)
			{
				ReportControl.QVDoc.Reload();
			}

			//method would have broken by now if the document could not be opened, this should probably be in a try-catch
			lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Qlik document opened.");
			Console.WriteLine("Qlik document opened.");
		}

		/// <summary>
		/// Opens the Word application and document associated with the global variable "wordPath"
		/// </summary>
		private void openWordDocument()
		{
			//need the Microsoft.Office.Interop.Word reference in order for this to work properly
			ReportControl.WordApp = new Word.Application();
			ReportControl.WordApp.Visible = false;
			ReportControl.WordDoc = ReportControl.WordApp.Documents.Open(ReportControl.wordPath);

			//method would have broken by now if the document could not be opened, this should probably be in a try-catch
			lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Word document opened.");
			Console.WriteLine("Word Document opened.");
		}

		/// <summary>
		/// Checking that the paths set by the file browser dialogs are still valid when the user clicks the Generate button.
		/// </summary>
		private bool validatePaths()
		{
			//return false if the qlik or word paths are not valid, true if they are valid
			if (File.Exists(txtWordPath.Text) && File.Exists(txtQlikPath.Text))
			{
				//validate the file types just in case something went wrong with the OpenFileDialog
				if (txtQlikPath.Text.EndsWith(".qvw") && (txtWordPath.Text.EndsWith(".docx") || txtWordPath.Text.EndsWith(".doc")))
				{
					return true;
				}
				else
				{
					MessageBox.Show("Improper file types", "Invalid Types", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return false;
				}
			}
			else
			{
				MessageBox.Show("Document paths could not be validated", "Invalid Paths", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
		}

		/// <summary>
		/// Opening file browser to search for .qvw files
		/// </summary>
		private void btnQlikBrowse_Click(object sender, EventArgs e)
		{
			//open file browser and limit results to .qvw files (Qlik Sense files will have a different file type)
			OpenFileDialog qlik = new OpenFileDialog();
			qlik.Filter = "QlikView Files (*.qvw)|*.qvw";
			qlik.Multiselect = false;

			if (qlik.ShowDialog() == DialogResult.OK)
			{
				//if an acceptable file is chosen, load the path into the uneditable text box
				txtQlikPath.Text = qlik.FileName;
			}
		}

		/// <summary>
		/// Opening file browser to search for .doc or .docx files
		/// </summary>
		private void btnWordBrowse_Click(object sender, EventArgs e)
		{
			//open file browser and limit results to .doc and .docx
			//this program has only been tested with Microsoft Word 2010 and 2013
			OpenFileDialog word = new OpenFileDialog();
			word.Filter = "Word Files (*.doc;*.docx)|*.doc;*.docx";
			word.Multiselect = false;

			if (word.ShowDialog() == DialogResult.OK)
			{
				//if an acceptable file is chosen, load the p   ath into the uneditable text box
				txtWordPath.Text = word.FileName;
			}
		}

		/// <summary>
		/// Closes QlikView document and aapplication cleanly
		/// </summary>
		private void closeQlikDocument()
		{
			//not sure the if statements are completely necessary, but this gets the job done
			if (ReportControl.QVDoc != null) ReportControl.QVDoc.CloseDoc();
			if (ReportControl.QVApp != null) ReportControl.QVApp.Quit();
			ReportControl.QVApp = null;
			ReportControl.QVDoc = null;

			lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Qlik document closed.");
			Console.WriteLine("Qlik document closed.");
		}

		/// <summary>
		/// Closes Word document and application cleanly
		/// </summary>
		private void closeWordDocument()
		{
			string savePath = "";

			//not sure the if statements are completely necessary, but this gets the job done
			if (ReportControl.WordDoc != null)
			{
				savePath = Path.GetFullPath(ReportControl.wordPath).Replace(Path.GetExtension(ReportControl.wordPath), "-v"+DateTime.Now.ToString("yyyyMMddHHmm")) + Path.GetExtension(ReportControl.wordPath);
				Console.WriteLine(savePath);
				lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Saving completed print " + savePath);
				ReportControl.WordDoc.SaveAs2(savePath);
				ReportControl.WordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
			}
			ReportControl.WordApp.Quit();
			ReportControl.WordApp = null;
			ReportControl.WordDoc = null;

			if (cbxOpenWord.Checked)
			{
				var applicationWord = new Word.Application();
				applicationWord.Visible = true;
				applicationWord.Documents.Open(savePath);
				lstLog.TopIndex = lstLog.Items.Count - 1;
				lstLog.Items.Add("Opening word document.");
			}
			else
			{
				lstLog.TopIndex = lstLog.Items.Count - 1;
				lstLog.Items.Add("Word document closed.");
				Console.WriteLine("Word document closed.");
			}
		}

		/// <summary>
		/// Cleans up the QlikView and Word documents and applications
		/// </summary>
		private void exitWithGrace()
		{

			//call specialized closing methods
			closeQlikDocument();
			closeWordDocument();

			lstLog.TopIndex = lstLog.Items.Count - 1; lstLog.Items.Add("Finished cleaning up.");
			Console.WriteLine("Finished cleaning up.");
		}

		/// <summary>
		/// Save the quick reference variable to dictionary
		/// </summary>
		private void btnSaveRef_Click(object sender, EventArgs e)
		{
			string RefName, RefID;
			char[] trimchar = { '<', '>', '!' };

			if (txtRefName.Text != "" && txtRefID.Text != "")
			{
				if (txtRefName.Text.Contains("-") | txtRefID.Text.Contains("-"))
				{
					MessageBox.Show("Your inputs may not contain a hyphen (-).", "Invalid Character", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					RefName = txtRefName.Text.Trim(trimchar);
					RefID = txtRefID.Text.Trim(trimchar);
					try
					{
						QuickRefVars.Add(RefName, Tuple.Create<string, string>(RefID, ""));
						lstQuickRefVars.Items.Add(RefName + " - " + RefID);
						txtRefName.Clear();
						txtRefID.Clear();
					}
					catch
					{
						MessageBox.Show("Could not add item. You may already have added an item with the same reference text.", "Unsuccessful", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			else
			{
				MessageBox.Show("You must provide a value for both fields.", "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void btnRemove_Click(object sender, EventArgs e)
		{
			if (lstQuickRefVars.SelectedItem != null)
			{
				string listItem = lstQuickRefVars.GetItemText(lstQuickRefVars.SelectedItem);
				int index = listItem.IndexOf("-");
				listItem = listItem.Remove(index - 1);

				lstQuickRefVars.Items.Clear();
				QuickRefVars.Remove(listItem);

				foreach (KeyValuePair<string, Tuple<string, string>> entry in QuickRefVars)
				{
					lstQuickRefVars.Items.Add(entry.Key + " - " + entry.Value.Item1);
				}
			}
		}
	}

	internal static class ReportControl
	{
		//some global variables for Word, Qlik, and the list of static selections
		public static string wordPath { get; set; }
		public static Word.Application WordApp { get; set; }
		public static Word.Document WordDoc { get; set; }

		public static string qlikPath { get; set; }
		public static QlikView.Application QVApp { get; set; }
		public static Doc QVDoc { get; set; }

		public static List<Tag> staticSelections { get; set; }
		public static System.Windows.Forms.ListBox listLog { get; set; }
	}

	//tag class, with a field and selection
	//a list of tags represents a single {} bracketed object
	internal class Tag
	{
		public string Field { get; set; }
		public string Selection { get; set; }

		public Tag(string field, string selection)
		{
			this.Field = field;
			this.Selection = selection;
		}

		public Tag(Tuple<string, string> selectionTag)
		{
			this.Field = selectionTag.Item1;
			this.Selection = selectionTag.Item2;
		}


		//combine two tags with the same field, differing fields returns the original tag
		public static Tag combine(Tag t1, Tag t2)
		{

			if (t1.Field == t2.Field)
			{

				string[] t1Sels = t1.Selection.Split('|');
				string[] t2Sels = t2.Selection.Split('|');
				IEnumerable<string> bothSels = t1Sels.Concat(t2Sels);
				bothSels.Distinct();
				string s = "";
				foreach (string i in bothSels.Distinct())
				{
					s += i + "|";
				}
				Tag tRet = new Tag(t1.Field, s.TrimEnd('|'));
				return tRet;
			}
			else
			{
				return t1;
			}
		}

		//returns the string of the format {Field,Selection}
		public override string ToString()
		{
			return "{'" + this.Field.Replace("','", "#%1234") + "','" + this.Selection.Replace("','", "#%1234") + "'}";
		}

		//tries to aggregate all tags of the same type in a list
		public static List<Tag> aggTags(List<Tag> tags)
		{
			List<Tag> matches = new List<Tag>();
			for (int i = 0; i < tags.Count; i++)
			{
				matches = tags.FindAll(z => z.Field == tags[i].Field);
				if (matches != null && matches.Count > 1)
				{
					Tag newTag = tags[i];
					tags.RemoveAll(z => z.Field == tags[i].Field);
					foreach (Tag z in matches)
					{
						newTag = combine(newTag, z);
					}
					tags.Add(newTag);
				}
			}
			return tags;
		}

		//returns a list of tags as a single string
		public static string listTagsToString(List<Tag> tags)
		{
			if (tags == null || tags.Count <= 0)
			{
				return "";
			}
			else
			{
				string retTag = "{";
				foreach (Tag i in tags)
				{
					retTag += "'" + i.Field.Replace("','", "#%1234") + "','" + i.Selection.Replace("','", "#%1234") + "',";
				}
				retTag = retTag.TrimEnd(",".ToCharArray());
				retTag += "}";
				return retTag;
			}
		}

		public static List<Tag> addTag(List<Tag> tags, Tag newTag)
		{
			if (tags == null)
			{
				tags = new List<Tag>();
				tags.Add(newTag);
				return tags;
			}
			tags.Add(newTag);
			return aggTags(tags);
		}

		//rework of interpretSelectionTag
		public static List<Tag> interpretSelectionTag(string selectionTag)
		{
			selectionTag = selectionTag.Replace('’', '\'');
			selectionTag = selectionTag.Replace('‘', '\'');
			Console.WriteLine("Selection tag to interpret: {0}", selectionTag);
			List<Tag> selections = new List<Tag>();
			if (selectionTag != "")
			{
				if (Equals(selectionTag[0].ToString(), "{") && Equals(selectionTag[selectionTag.Length - 1].ToString(), "}"))
				{
					Regex rgx = new Regex(@"','");
					int commaCount = rgx.Matches(selectionTag).Count;

					if (commaCount % 2 != 0) //properly formatted tags will always have an odd number of commas
					{

						//error checking complete, now manipulate strings to make list
						while (selectionTag.Contains(@"','"))
						{
							int commaIndex;
							string fieldName, fieldValue;

							commaIndex = selectionTag.IndexOf(@"','"); //get index of first comma
							fieldName = selectionTag.Substring(2, commaIndex - 2);

							if (rgx.Matches(selectionTag).Count == 1)
							{
								selectionTag = selectionTag.Remove(1, commaIndex + 2);
								fieldValue = selectionTag.Substring(1, selectionTag.Length - 3);
								selectionTag = selectionTag.Remove(1, fieldValue.Length + 1);
								Console.WriteLine(selectionTag);
							}
							else
							{
								selectionTag = selectionTag.Remove(1, fieldName.Length + 4);
								commaIndex = selectionTag.IndexOf(@"','");
								fieldValue = selectionTag.Substring(1, commaIndex - 1);
								selectionTag = selectionTag.Remove(1, fieldValue.Length + 2);
							}

							selections.Add(new Tag(fieldName.Trim().Replace("#%1234", "','"), fieldValue.Trim().Replace("#%1234", "','")));
						}

						return selections;
					}
					else
					{
						Console.WriteLine("Selection tag invalid, incorrect number of arguments: {0}", selectionTag);
						return null;
					}
				}
				else
				{
					Console.WriteLine("Selection tag invalid, improper brace structure: {0}", selectionTag);
					return null;
				}
			}
			else
			{
				Console.WriteLine("Selection tag was null");
				return selections;
			}
		}
	}

	//chart class, with chart name and selection tags
	//used to represents a <> bracketed object
	internal class Chart
	{
		public string Name { get; set; }
		public List<Tag> Selections { get; set; }
		public string attributes { get; set; }

		public Chart(string name, List<Tag> selections)
		{
			this.Name = name;
			this.Selections = selections;
		}

		//reads in chart of format [chartname{Field1,Selection1,Field2,Selection2...}]
		public Chart(string cString)
		{
			if (!cString.Contains('{') && !cString.Contains('?'))
			{
				this.Name = cString.Trim("<> ".ToCharArray());
				this.Selections = null;
				this.attributes = null;
			}
			else if (cString.Contains('{') && !cString.Contains('?'))
			{
				Regex rgx = new Regex("{.*}", RegexOptions.IgnorePatternWhitespace);
				string findSelections = rgx.Match(cString).Value;
				this.Selections = Tag.interpretSelectionTag(findSelections);
				rgx = new Regex("<.*{", RegexOptions.IgnorePatternWhitespace);
				findSelections = rgx.Match(cString).Value;
				this.Name = findSelections.Trim("<{ ".ToCharArray());
				this.attributes = null;
			}
			else
			{
				string[] split = cString.Split('?');
				this.attributes = "?" + split[1].Trim(" >".ToCharArray());

				Regex rgx = new Regex("{.*}", RegexOptions.IgnorePatternWhitespace);
				string findSelections = rgx.Match(split[0]).Value;
				this.Selections = Tag.interpretSelectionTag(findSelections);
				rgx = new Regex("<.*{", RegexOptions.IgnorePatternWhitespace);
				findSelections = rgx.Match(split[0]).Value;
				this.Name = findSelections.Trim("<{ ".ToCharArray());
			}
		}

		public override string ToString()
		{
			return "<" + this.Name + Tag.listTagsToString(this.Selections) + this.attributes + ">";
		}

		public Chart AddSelectionTag(Tag selection)
		{
			this.Selections = Tag.addTag(this.Selections, selection);
			return this;
		}

		public Chart AddSelectionTag(List<Tag> selection)
		{
			foreach (Tag t in selection)
			{
				this.AddSelectionTag(t);
			}
			return this;
		}
	}
}