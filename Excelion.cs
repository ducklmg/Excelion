using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Excelion
{
	public class Excelion
	{
		string[] mLanguages;
		Dictionary<string, string[]> mTable;

		int mActiveLang = 0;
		string mEmptyStringValue = null;

		static public Excelion Load(string filePath)
		{
			return FromJson(File.ReadAllText(filePath, Encoding.UTF8));
		}

		static public Excelion Load(byte[] jsonBytes)
		{
			return FromJson(Encoding.UTF8.GetString(jsonBytes));
		}

		public void Save(string filePath)
		{
			File.WriteAllText(filePath, ToJson(), Encoding.UTF8);
		}

		static public Excelion FromJson(string json)
		{
			try
			{
				var data = MiniJSON.Json.Deserialize(json) as Dictionary<string, object>;

				var languages = Util.ToStringArray(data["Languages"]);

				var table = new Dictionary<string, string[]>();

				foreach( var item in (Dictionary<string, object>)data["Table"] )
				{
					table[item.Key] = Util.ToStringArray(item.Value);
				}

				return new Excelion()
				{
					mLanguages = languages,
					mTable = table
				};
			}
			catch( Exception )
			{
				return null;
			}
		}

		public string ToJson()
		{
			var data = new Dictionary<string, object>();
			data["Languages"] = mLanguages;
			data["Table"] = mTable;

			return MiniJSON.Json.Serialize(data);
		}

		public int GetLanguageIndex(string lang)
		{
			for( int i = 0; i < mLanguages.Length; i++ )
				if( lang.Equals(mLanguages[i], StringComparison.OrdinalIgnoreCase) )
					return i;

			return 0;
		}

		public void SetActiveLanguage(string lang)
		{
			mActiveLang = GetLanguageIndex(lang);
		}

		public string GetActiveLanguage()
		{
			return mLanguages[mActiveLang];
		}

		public string GetString(string id)
		{
			return GetString(id, mActiveLang);
		}

		public string GetString(string id, string lang)
		{
			return GetString(id, GetLanguageIndex(lang));
		}

		public string GetString(string id, int langIdx)
		{
			string[] values;
			if( mTable.TryGetValue(id, out values) )
			{
				if( 0 <= langIdx && langIdx < values.Length )
				{
					return values[langIdx];
				}
			}

			return mEmptyStringValue;
		}

		public void SetEmptyStringValue(string emptyStr)
		{
			mEmptyStringValue = emptyStr;
		}

		internal void SetData(string[] languages, Dictionary<string, string[]> table)
		{
			mLanguages = languages;
			mTable = table;
		}

		class CsvParser
		{
			/* RFC 4180 : "Common Format and MIME Type for Comma-Separated Values (CSV) Files"
				file = [header CRLF] record *(CRLF record) [CRLF]
				header = name *(COMMA name)
				record = field *(COMMA field)
				name = field
				field = (escaped / non-escaped)
				escaped = DQUOTE *(TEXTDATA / COMMA / CR / LF / 2DQUOTE) DQUOTE
				non-escaped = *TEXTDATA
				COMMA = %x2C
				CR = %x0D
				DQUOTE =  %x22
				LF = %x0A
				CRLF = CR LF
				TEXTDATA =  %x20-21 / %x23-2B / %x2D-7E
			*/
			string mText;
			int mPos;
			int mLength;

			StringBuilder mStrBuilder;
			List<string> mFields;

			public CsvParser(string csv)
			{
				mText = csv;
				mPos = 0;
				mLength = mText.Length;

				mStrBuilder = new StringBuilder(256);
				mFields = new List<string>(8);
			}

			public IEnumerable<string[]> Records
			{
				get
				{
					var record = GetRecord();
					if( record == null )
						yield break;

					yield return record;
				}
			}

			public string[] GetRecord()
			{
				if( mPos == mLength )
					return null;

				mFields.Clear();

				while( true )
				{
					string field = GetField();
					if( field == null )
						break;

					mFields.Add(field);
				}

				return mFields.ToArray();
			}

			string GetField()
			{
				if( mPos == mLength )
					return null;

				if( NextIsCRLF )
				{
					mPos += 2;
					return null;
				}

				if( mText[mPos]==QUOTE )
				{
					return GetQuoteString();
				}
				else
				{
					return GetString();
				}
			}

			private string GetString()
			{
				mStrBuilder.Clear();

				while( mPos < mLength )
				{
					char ch = mText[mPos];
					if( ch == '\r' || ch == '\n' )
						break;

					mPos++;

					if( ch == COMMA )
						break;

					mStrBuilder.Append(ch);
				}

				return mStrBuilder.ToString();
			}

			private string GetQuoteString()
			{
				mStrBuilder.Clear();

				while( mPos < mLength )
				{
					char ch = mText[mPos++];
					if( ch == QUOTE )
					{
						bool doubleQuote = mPos < mLength && mText[mPos + 1] == QUOTE;
						if( doubleQuote )
							mPos++;
						else
							break;
					}

					mStrBuilder.Append(ch);
				}

				if( mPos < mLength && mText[mPos] == COMMA )
					mPos++;

				return mStrBuilder.ToString();
			}

			const char COMMA = ',';
			const char QUOTE = '"';

			bool NextIsCRLF
			{
				get
				{
					return mPos + 1 < mLength && mText[mPos] == '\r' && mText[mPos + 1] == '\n';
				}
			}
		}
	}
}
