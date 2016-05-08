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
			enum Token
			{
				Character,
				EOF,
				NewLine,
				Comma,
				Quote,
			}

			StringBuilder mBuilder;
            string mText;
            int mPos;
			char mChar;
			bool mEndOfLine;

			public bool EOF { get { return mPos == mText.Length; } }

			Token NextChar()
			{
				if( EOF )
					return Token.EOF;

				char ch = mText[mPos++];

				switch( ch )
				{
					case '\n':
						return Token.NewLine;
					case '\r':
						return NextChar();
					case ',':
						return Token.Comma;
					case '"':
						return Token.Quote;
					case '\\':
						if( EOF )
							return Token.EOF;

						char nextCh = mText[mPos];
						switch( nextCh )
						{
							case 'b': mChar = '\b'; break;
							case 'f': mChar = '\f'; break;
							case 'n': mChar = '\n'; break;
							case 'r': mChar = '\r'; break;
							case 't': mChar = '\t'; break;

							case '/': mChar = '/'; break;
							case '"': mChar = '"'; break;
							case '\'': mChar = '\''; break;
							case '\\': mChar = '\\'; break;

							// unrecognized sequence. ignore backslash
							default:
								return NextChar();
						}

						mPos++;
						return Token.Character;

					default:
						mChar = ch;
						return Token.Character;
				}
			}

			string FetchValue()
			{
				mBuilder.Clear();



				bool quoteMode = mText[mPos] == '"';




				Token token = NextChar();
				if( token == Token.Quote )
				{
					quoteMode = true;
					token = NextChar();
				}

				while( true )
				{
					switch( token )
					{
						case Token.Character:
							mBuilder.Append(mChar);
							break;

						case Token.Quote:
							if( quoteMode )
							{
								string value = mBuilder.ToString();

								var comma = NextChar();
								if( comma != Token.Comma )
									throw new FormatException();

								return value;
							}
							else
							{
								mBuilder.Append(mChar);
							}
							break;

						case Token.NewLine:
							mEndOfLine = true;
							return mBuilder.ToString();

						case Token.EOF:
							return mBuilder.ToString();

						case Token.Comma:
							if( quoteMode )
							{
								mBuilder.Append(mChar);
							}
							else
							{
								return mBuilder.ToString();
							}
							break;
					}
				}
			}

			int findChar(int startPos, char ch)
			{
				for( int i = startPos; i < mText.Length; i++ )
				{
					if( mText[i] == ch )
						return i;
				}

				return -1;
			}

			void str()
			{
				if( mText[mPos] == '"' )
				{
					int start = mPos + 1;
					int end = findChar(start, '"');
					if( end == -1 )
						;

					if( end + 1 < mText.Length && mText[end + 1] == '"' )
						;
				}
				else
				{
					int start = mPos;
					int end = findChar(start, ',');

				}
			}
        }
	}
}
