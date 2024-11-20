/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures
{
    internal class StructureKeyNamesCache
    {
        private int _nextId = 0;
        private readonly Dictionary<string, int> _words = new Dictionary<string, int>();
        private readonly Dictionary<int, string> _wordsById = new Dictionary<int, string>();

        public int GetId(string word)
        {
            if(!_words.ContainsKey(word))
            {
                var id = ++_nextId;
                _words[word] = id;
                _wordsById[id] = word;
                return id;
            }
            return _words[word];
        }

        public List<int> GetIds(IEnumerable<string> words)
        {
            var result = new List<int>();
            foreach(var word in words)
            {
                var id = GetId(word);
                result.Add(id);
            }
            return result;
        }

        public string GetWord(int id)
        {
            if(!_wordsById.ContainsKey(id))
            {
                return null;
            }
            return _wordsById[id];
        }
    }
}
