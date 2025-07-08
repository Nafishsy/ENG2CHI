# Standard library imports
from pathlib import Path
from datetime import datetime
import gc
import os
import time
import shutil
import logging
import zipfile
import re  # Add re import for regular expressions

# Third-party imports
import win32com.client
import pythoncom
import win32gui
import win32process
import win32con
import win32api
import psutil
from lxml import etree

# Logging setup
def init_logging(log_file_path):
    """Initialize logging to both file and console with timestamps"""
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    
    # Create timestamped log file path
    log_path = Path(log_file_path).parent / f'sync_changes_{timestamp}.log'
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Create file handler
    fh = logging.FileHandler(str(log_path), 'w', 'utf-8')
    fh.setLevel(logging.DEBUG)  # Log everything to file
    fh.setFormatter(formatter)
    
    # Create console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)  # Less verbose on console
    ch.setFormatter(formatter)
    
    # Get the root logger and set its level
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)  # Capture all levels
    
    # Remove any existing handlers
    for handler in root.handlers[:]:
        root.removeHandler(handler)
    
    # Add the new handlers
    root.addHandler(fh)
    root.addHandler(ch)
    
    return log_path

def find_matching_documents():
    """Find matching English and Chinese document pairs with improved name matching"""
    try:
        # Find English documents with track changes
        english_docs = []
        for doc in Path('.').glob('*.docx'):
            if ('[Track Changes]' in doc.name and 
                not doc.name.startswith('~$')):  # Skip temp files
                english_docs.append(doc)
                
        if not english_docs:
            logging.error("No English documents with track changes found")
            return []
            
        logging.info(f"Found {len(english_docs)} English document(s) with track changes")
        pairs = []
        
        # Process each English document
        for eng_doc in english_docs:
            try:
                # Get base name without markers
                name = eng_doc.stem
                name = name.replace('[Track Changes]', '').strip()
                if '_E' in name:
                    name = name.split('_E')[0].strip()
                    
                logging.info(f"Looking for Chinese match for: {name}")
                
                # Find matching Chinese documents
                best_match = None
                best_score = 0
                
                for doc in Path('.').glob('*.docx'):
                    if (doc.name.startswith('~$') or  # Skip temp files
                        '[Track Changes]' in doc.name or  # Skip English files
                        doc == eng_doc):  # Skip self
                        continue
                        
                    # Score the match
                    score = 0
                    if name in doc.name:
                        score += 5
                    if '_C' in doc.name:
                        score += 3
                    if name == doc.stem.replace('_C', '').strip():
                        score += 10
                        
                    if score > best_score:
                        best_score = score
                        best_match = doc
                
                if best_match and best_score >= 5:
                    logging.info(f"Found Chinese match: {best_match.name}")
                    pairs.append((eng_doc, best_match))
                else:
                    logging.warning(f"No matching Chinese document found for {eng_doc.name}")
                    
            except Exception as e:
                logging.error(f"Error matching documents: {str(e)}")
                continue
                
        logging.info(f"Found {len(pairs)} document pair(s)")
        return pairs
        
    except Exception as e:
        logging.error(f"Error finding document pairs: {str(e)}")
        return []

class DocumentTracker:
    def __init__(self):
        # Cache for paragraph analysis
        self._para_cache = {}
        self._cache_hits = 0
        self._cache_misses = 0
        self.word_app = None
        self._ensure_word_closed()
        self._init_word()
        
    def __enter__(self):
        """Context manager entry"""
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit with cleanup"""
        self._cleanup_word()
        
    def _ensure_word_closed(self):
        """Ensure all Word instances are closed before starting"""
        try:
            # Try to close Word windows gracefully first
            word_windows = []
            def enum_window_callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if "Microsoft Word" in window_text or ".docx" in window_text:
                        try:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            word_windows.append((hwnd, pid))
                        except:
                            pass
            win32gui.EnumWindows(enum_window_callback, None)
            
            for hwnd, _ in word_windows:
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except:
                    pass
                    
            # Give windows time to close
            time.sleep(2)
            
            # Force close any remaining Word processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'WINWORD.EXE' in proc.name().upper():
                        proc.kill()
                except:
                    pass
                    
            # Final cleanup
            pythoncom.CoInitialize()
            try:
                existing_app = win32com.client.GetActiveObject('Word.Application')
                existing_app.Quit()
            except:
                pass
            finally:
                pythoncom.CoUninitialize()
                
        except Exception as e:
            logging.warning(f"Error during Word cleanup: {str(e)}")
            
    def _init_word(self):
        """Initialize Word application with robust error handling"""
        if self.word_app is None:
            try:
                pythoncom.CoInitialize()
                self.word_app = win32com.client.dynamic.Dispatch('Word.Application')
                self.word_app.Visible = False
                self.word_app.DisplayAlerts = False
                logging.info("Word application initialized successfully")
            except Exception as e:
                logging.error(f"Failed to initialize Word application: {str(e)}")
                self._cleanup_word()
                raise
                
    def _cleanup_word(self):
        """Clean up Word application and COM resources"""
        try:
            if self.word_app:
                try:
                    # Close all open documents
                    if hasattr(self.word_app, 'Documents'):
                        for i in range(self.word_app.Documents.Count, 0, -1):
                            try:
                                doc = self.word_app.Documents.Item(i)
                                doc.Close(SaveChanges=False)
                            except:
                                pass
                except:
                    pass
                
                try:
                    self.word_app.Quit()
                except:
                    pass
                
                self.word_app = None
                
            pythoncom.CoUninitialize()
            logging.info("Word application closed")
            
        except Exception as e:
            logging.error(f"Error during Word cleanup: {str(e)}")
            # Make sure we don't leave hanging processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'WINWORD.EXE' in proc.name().upper():
                        proc.kill()
                except:
                    pass
                    
    def process_document_pair(self, english_doc, chinese_doc):
        """Process a pair of documents with improved error handling"""
        backup_path = None
        doc = None
        try:
            # Create backup of Chinese document
            backup_path = f"{chinese_doc}.backup"
            shutil.copy2(chinese_doc, backup_path)
            logging.info(f"Created backup: {backup_path}")
            
            # Extract changes from English document
            changes = []
            try:
                logging.info(f"Opening English document: {english_doc}")
                doc = self._open_document_with_retry(english_doc)
                
                if not doc:
                    raise Exception("Failed to open English document")
                    
                logging.info("Successfully opened English document")
                changes = self._extract_changes_from_doc(doc)
                
                if not changes:
                    raise Exception("No changes found in English document")
                    
                logging.info(f"Successfully extracted {len(changes)} changes")
                
            except Exception as e:
                logging.error(f"Error processing English document: {str(e)}")
                raise
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
                        
            # Apply changes to Chinese document
            success = self.apply_changes_to_chinese(chinese_doc, changes)
            if not success:
                raise Exception("Failed to apply changes to Chinese document")
                
            logging.info(f"Successfully processed document pair")
            return True
            
        except Exception as e:
            logging.error(f"Error processing document pair: {str(e)}")
            if backup_path:
                self._restore_from_backup(chinese_doc, backup_path)
            return False
            
    def _get_context_with_paragraphs(self, doc, revision):
        """Get context including paragraph structure and formatting with better error handling"""
        try:
            rev_range = revision.Range
            context = {
                'before': '',
                'current': '',
                'after': '',
                'structure': {}
            }
            
            try:
                # Get current paragraph
                current_para = rev_range.Paragraphs.First
                if not current_para:
                    return None
                    
                context['current'] = current_para.Range.Text.strip()
                
                # Get structural information
                try:
                    context['structure'] = {
                        'style': current_para.Style.NameLocal if hasattr(current_para.Style, 'NameLocal') else None,
                        'level': current_para.OutlineLevel,
                        'alignment': current_para.Alignment,
                        'line_spacing': current_para.LineSpacing,
                        'first_line_indent': current_para.FirstLineIndent
                    }
                except Exception as e:
                    logging.debug(f"Could not get some paragraph properties: {str(e)}")
                
                # Get surrounding paragraphs
                try:
                    # Get previous paragraph
                    prev_range = rev_range.Paragraphs.First.Range.Paragraphs.First.Previous
                    if prev_range:
                        context['before'] = prev_range.Text.strip()
                except:
                    pass
                    
                try:
                    # Get next paragraph
                    next_range = rev_range.Paragraphs.First.Range.Paragraphs.First.Next
                    if next_range:
                        context['after'] = next_range.Text.strip()
                except:
                    pass
                
                return context
                
            except Exception as e:
                logging.warning(f"Error getting paragraph context: {str(e)}")
                # Fallback to basic context if paragraph access fails
                context['current'] = rev_range.Text.strip()
                try:
                    context['before'] = rev_range.Previous(50).Text.strip()  # Get 50 chars before
                except:
                    pass
                try:
                    context['after'] = rev_range.Next(50).Text.strip()  # Get 50 chars after
                except:
                    pass
                return context
                
        except Exception as e:
            logging.error(f"Error getting context: {str(e)}")
            return None
            
    def _extract_changes_from_doc(self, doc):
        """Extract changes from an open Word document"""
        try:
            changes = []
            revision_count = doc.Revisions.Count
            logging.info(f"Found {revision_count} revisions in document")
            
            if revision_count == 0:
                logging.warning("No revisions found in document")
                return changes
                
            for i in range(revision_count):
                try:
                    revision = doc.Revisions.Item(i + 1)
                    if not revision:
                        logging.warning(f"Could not access revision {i + 1}")
                        continue
                        
                    # Get revision properties
                    rev_type = revision.Type
                    rev_text = revision.Range.Text
                    
                    # Skip empty or non-text changes
                    if not rev_text or rev_type not in [1, 2]:  # 1=Insert, 2=Delete
                        logging.debug(f"Skipping revision {i + 1}: empty or non-text change")
                        continue
                        
                    logging.debug(f"Processing revision {i + 1}: Type={rev_type}, Text={rev_text[:50]}...")
                    
                    # Get context for the change
                    context = self._get_context_with_paragraphs(doc, revision)
                    if not context:
                        logging.warning(f"Could not get context for revision {i + 1}")
                        continue
                        
                    changes.append({
                        'revision': revision,
                        'text': rev_text,
                        'type': rev_type,
                        'context': context
                    })
                    logging.debug(f"Successfully processed revision {i + 1}")
                    
                except Exception as e:
                    logging.error(f"Error processing revision {i + 1}: {str(e)}")
                    continue
                    
            logging.info(f"Successfully extracted {len(changes)} changes from {revision_count} revisions")            return changes
            
        except Exception as e:
            logging.error(f"Error extracting changes: {str(e)}")
            return []
            
    def _apply_single_change(self, doc, change, retry=False):
        """Apply a single change to a document with number preservation"""
        try:
            revision = change['revision']
            rev_type = change['type']
            rev_text = change['text']
            context = change['context']
            
            # Find matching position in target document
            target_range = self._find_match_with_context(doc, rev_text, context, retry=retry)
            if not target_range:
                logging.warning(f"Could not find match for change: {rev_text[:50]}...")
                return False
            
            if rev_type == 1:  # Insert
                # For insertions, preserve numbers from the original text
                original_text = target_range.Text.strip()
                new_text = rev_text
                
                # Extract numbers from both texts
                original_chunks = self._extract_number_chunks(original_text)
                new_chunks = self._extract_number_chunks(new_text)
                
                # Replace numbers in the new text with original numbers if they exist
                original_numbers = [chunk for chunk in original_chunks if chunk.replace('.', '').replace('%', '').isdigit()]
                new_numbers = [chunk for chunk in new_chunks if chunk.replace('.', '').replace('%', '').isdigit()]
                
                if len(original_numbers) == len(new_numbers):
                    for orig, new in zip(original_numbers, new_numbers):
                        new_text = new_text.replace(new, orig, 1)
                
                target_range.Text = new_text
            elif rev_type == 2:  # Delete
                target_range.Delete()
            
            return True
            
        except Exception as e:
            logging.error(f"Error applying change: {str(e)}")
            return False

    def _extract_number_chunks(self, text):
        """Extract numbers and their surrounding text from a string"""
        # Match numbers including decimals, percentages, and those with surrounding punctuation
        pattern = r'(\d+(?:\.\d+)?%?)|([^\d]+)'
        chunks = re.findall(pattern, text)
        # Flatten and filter empty strings
        return [chunk for tuple_chunk in chunks for chunk in tuple_chunk if chunk]    def _apply_single_change(self, doc, change, retry=False):
        """Apply a single change to a document with number preservation"""
        try:
            revision = change['revision']
            rev_type = change['type']
            rev_text = change['text']
            context = change['context']
            
            # Find matching position in target document
            target_range = self._find_match_with_context(doc, rev_text, context, retry=retry)
            if not target_range:
                logging.warning(f"Could not find match for change: {rev_text[:50]}...")
                return False
            
            if rev_type == 1:  # Insert
                # For insertions, preserve numbers from the original text
                original_text = target_range.Text.strip()
                new_text = rev_text
                
                # Extract numbers from both texts
                original_chunks = self._extract_number_chunks(original_text)
                new_chunks = self._extract_number_chunks(new_text)
                
                # Replace numbers in the new text with original numbers if they exist
                original_numbers = [chunk for chunk in original_chunks if chunk.replace('.', '').replace('%', '').isdigit()]
                new_numbers = [chunk for chunk in new_chunks if chunk.replace('.', '').replace('%', '').isdigit()]
                
                if len(original_numbers) == len(new_numbers):
                    for orig, new in zip(original_numbers, new_numbers):
                        new_text = new_text.replace(new, orig, 1)
                
                target_range.Text = new_text
            elif rev_type == 2:  # Delete
                target_range.Delete()
            
            return True
            
        except Exception as e:
            logging.error(f"Error applying change: {str(e)}")
            return False
            
    def _calculate_text_similarity(self, text, analysis):
        """Calculate text similarity using multi-stage matching including AI translation comparison"""
        try:
            # Stage 1: Handle very short text with more lenient matching
            text = text.strip()
            analysis_text = analysis['text'].strip()
            is_very_short = len(text) <= 5 or len(analysis_text) <= 5

            # Stage 2: Enhanced number pattern matching
            number_pattern = r'(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?%?(?:\s*(?:USD|HKD|CNY|EUR|GBP|JPY))?'
            text_numbers = set(re.findall(number_pattern, text))
            analysis_numbers = set(re.findall(number_pattern, analysis_text))

            # Stage 3: Calculate length ratio score
            len_ratio = min(len(text), len(analysis_text)) / max(len(text), len(analysis_text)) if text and analysis_text else 0
            
            # Stage 4: Determine content-based weights
            if text_numbers and analysis_numbers:
                number_weight = 0.5
                char_weight = 0.2
                seq_weight = 0.2
                len_weight = 0.1
            elif is_very_short:
                number_weight = 0.2
                char_weight = 0.4
                seq_weight = 0.3
                len_weight = 0.1
            else:
                number_weight = 0.2
                char_weight = 0.4
                seq_weight = 0.2
                len_weight = 0.2

            # Stage 5: Number similarity with exact matching
            if text_numbers and analysis_numbers:
                common_numbers = text_numbers.intersection(analysis_numbers)
                if not common_numbers and not is_very_short:
                    return 0
                number_score = len(common_numbers) / max(len(text_numbers), len(analysis_numbers))
            else:
                number_score = 0.5

            # Stage 6: Text cleaning and normalization
            text_clean = re.sub(number_pattern, '', text)
            analysis_clean = re.sub(number_pattern, '', analysis_text)
            
            # Normalize punctuation and whitespace
            punctuation_pattern = r'[，。！？：；（）、]'
            text_clean = re.sub(punctuation_pattern, ' ', text_clean)
            analysis_clean = re.sub(punctuation_pattern, ' ', analysis_clean)
            text_clean = ' '.join(text_clean.split())
            analysis_clean = ' '.join(analysis_clean.split())
            
            # Stage 7: Character-based similarity for Chinese text
            chars1 = set(text_clean)
            chars2 = set(analysis_clean)
            
            if chars1 and chars2:
                char_ratio = len(chars1.intersection(chars2)) / max(len(chars1), len(chars2))
                
                # Stage 8: Sequence matching for word order
                from difflib import SequenceMatcher
                seq_ratio = SequenceMatcher(None, text_clean, analysis_clean).ratio()
                
                # Stage 9: Calculate initial score
                base_score = (
                    number_score * number_weight +
                    char_ratio * char_weight +
                    seq_ratio * seq_weight +
                    len_ratio * len_weight
                )

                # Stage 10: Semantic similarity using AI translation patterns
                if base_score < 0.6 and len(text_clean) > 10:  # Only for longer text with low confidence
                    try:
                        # Common translation patterns for financial documents
                        patterns = {
                            r'投资目标|investment objective': 0.9,
                            r'风险因素|risk factors': 0.9,
                            r'基金经理|fund manager': 0.9,
                            r'净值|net asset value': 0.9,
                            r'收益率|yield': 0.9,
                            r'投资策略|investment strategy': 0.9,
                            r'管理费|management fee': 0.9,
                            r'申购|subscription': 0.9,
                            r'赎回|redemption': 0.9
                        }
                        
                        # Check for common translation patterns
                        pattern_score = 0
                        for pattern, confidence in patterns.items():
                            zh_en = pattern.split('|')
                            if len(zh_en) == 2:
                                if (zh_en[0] in text_clean and zh_en[1] in analysis_clean) or \
                                   (zh_en[1] in text_clean and zh_en[0] in analysis_clean):
                                    pattern_score = max(pattern_score, confidence)
                        
                        if pattern_score > 0:
                            base_score = max(base_score, pattern_score * 0.8)  # Boost score but with some uncertainty

                        # Stage 11: Synonym matching for financial terms
                        synonyms = {
                            '认购': ['申购', 'subscription'],
                            '买入': ['申购', 'purchase'],
                            '卖出': ['赎回', 'redemption'],
                            '回报': ['收益', 'return'],
                            '波动': ['波幅', 'volatility'],
                            '股息': ['分红', 'dividend']
                        }
                        
                        for term, syn_list in synonyms.items():
                            if term in text_clean:
                                for syn in syn_list:
                                    if syn in analysis_clean:
                                        base_score = max(base_score, 0.7)  # Good confidence for synonym matches
                    except Exception as e:
                        logging.debug(f"Error in semantic matching: {str(e)}")
                
                # Stage 12: Apply very short text bonus
                if is_very_short and base_score > 0.5:
                    base_score = min(1.0, base_score * 1.2)
                
                return base_score
            
            return 0.0

        except Exception as e:
            logging.debug(f"Error calculating text similarity: {str(e)}")
            return 0

    def _find_match_with_context(self, doc, text, context, timeout=30, retry=False):
        """Find matching text position using enhanced context matching and fuzzy search with timeout"""
        try:
            start_time = time.time()
            best_match = None
            best_score = 0
            
            # Calculate document sections for position-based filtering
            total_paras = doc.Paragraphs.Count
            target_pos = context.get('relative_position', 0.5)
            
            # Enhanced pattern matching for numbers and short text
            text = text.strip()
            is_number = bool(re.match(r'^\s*\d+(?:\.\d+)?%?\s*$', text))
            is_very_short = len(text) <= 5
            
            # Determine thresholds based on content type and retry status
            if retry:
                number_threshold = 0.7  # Lower threshold for retry
                short_threshold = 0.1
                text_threshold = 0.2
            else:
                number_threshold = 0.9
                short_threshold = 0.15
                text_threshold = 0.3

            # Special handling for numbers
            if is_number:
                # Extract and normalize number
                raw_number = re.search(r'\d+(?:\.\d+)?', text).group(0)
                number_formats = [
                    raw_number,  # Original format
                    f"{float(raw_number):.0f}",  # No decimal
                    f"{float(raw_number):.1f}",  # One decimal
                    f"{float(raw_number):.2f}",  # Two decimals
                ]
                if '%' in text:
                    number_formats.extend([f"{n}%" for n in number_formats])
                
                # Try each format in the context
                window_size = 10 if not retry else total_paras // 4
                target_index = int(target_pos * total_paras)
                start_index = max(1, target_index - window_size)
                end_index = min(total_paras, target_index + window_size)
                
                for i in range(start_index, end_index + 1):
                    para = doc.Paragraphs.Item(i)
                    para_text = para.Range.Text.strip()
                    
                    for num_format in number_formats:
                        if num_format in para_text:
                            context_score = self._calculate_context_score(para, context)
                            structure_score = self._calculate_structure_score(para, context)
                            position_score = self._calculate_position_score(para, context)
                            
                            total_score = (context_score * 0.5 + 
                                         structure_score * 0.3 + 
                                         position_score * 0.2)
                            
                            if total_score > best_score:
                                best_score = total_score
                                best_match = para.Range
                                
                            if total_score > number_threshold:
                                return para.Range
                
                # If still no match, try surrounding paragraphs
                if best_score > (number_threshold * 0.8):
                    return best_match
                    
            elif is_very_short:
                threshold = short_threshold
                window_size = total_paras // 2
            else:
                threshold = text_threshold
                window_size = total_paras // 2
            
            # Standard text matching
            target_index = int(target_pos * total_paras)
            start_index = max(1, target_index - window_size)
            end_index = min(total_paras, target_index + window_size)
            
            for i in range(start_index, end_index + 1):
                try:
                    if time.time() - start_time > timeout:
                        logging.warning("Match timeout exceeded, using best match found")
                        break
                        
                    para = doc.Paragraphs.Item(i)
                    para_text = para.Range.Text.strip()
                    
                    if not para_text:
                        continue

                    # Calculate comprehensive match score
                    text_score = self._calculate_text_similarity(text, {'text': para_text})
                    context_score = self._calculate_context_score(para, context)
                    structure_score = self._calculate_structure_score(para, context)
                    position_score = self._calculate_position_score(para, context)
                    
                    total_score = (text_score * 0.4 + 
                                 context_score * 0.3 + 
                                 structure_score * 0.2 + 
                                 position_score * 0.1)
                    
                    if total_score > best_score:
                        best_score = total_score
                        best_match = para.Range
                        logging.debug(f"New best match (score: {total_score:.2f})")
                        
                        if total_score > threshold:
                            return para.Range

                except Exception as e:
                    logging.debug(f"Error checking paragraph: {str(e)}")
                    continue

            if best_score >= (threshold * 0.8):  # Accept slightly lower scores on retry
                logging.info(f"Found acceptable match (score: {best_score:.2f}")
                return best_match

            logging.warning(f"No match found above threshold (best score: {best_score:.2f})")
            return None

        except Exception as e:
            logging.error(f"Error finding match: {str(e)}")
            return None

    def _open_document_with_retry(self, doc_path, max_retries=3, retry_delay=2):
        """Open a Word document with robust retry logic"""
        abs_path = str(Path(doc_path).resolve())
        last_error = None
        
        for attempt in range(max_retries):
            try:
                if not self.word_app:
                    self._init_word()
                    
                logging.info(f"Opening document (attempt {attempt + 1}): {abs_path}")
                doc = self.word_app.Documents.Open(
                    FileName=abs_path,
                    ReadOnly=False,
                    Visible=False,
                    NoEncodingDialog=True
                )
                
                if doc:
                    # Verify document is responsive
                    _ = doc.Name
                    logging.info(f"Successfully opened document on attempt {attempt + 1}")
                    return doc
                    
            except Exception as e:
                last_error = str(e)
                logging.error(f"Error opening document (attempt {attempt + 1}): {last_error}")
                
                if attempt < max_retries - 1:
                    # Try to recover
                    self._cleanup_word()
                    time.sleep(retry_delay * (attempt + 1))  # Exponential backoff
                    continue
                    
        raise Exception(f"Failed to open document after {max_retries} attempts: {last_error}")

    def _restore_from_backup(self, original_path, backup_path):
        """Restore document from backup"""
        try:
            if os.path.exists(backup_path):
                if os.path.exists(original_path):
                    os.remove(original_path)
                shutil.copy2(backup_path, original_path)
                logging.info(f"Restored from backup: {original_path}")
            else:
                logging.error(f"Backup file not found: {backup_path}")
        except Exception as e:
            logging.error(f"Error restoring from backup: {str(e)}")

    def apply_changes_to_chinese(self, doc_path, changes):
        """Apply changes to Chinese document with improved error handling and retry mechanism"""
        try:
            doc = self._open_document_with_retry(doc_path)
            if not doc:
                return False
            
            try:
                # Ensure test directory exists
                test_dir = Path('test')
                test_dir.mkdir(exist_ok=True)

                # Group changes by type for better processing
                number_changes = []
                text_changes = []
                for change in changes:
                    if bool(re.match(r'^\s*\d+(?:\.\d+)?%?\s*$', change['text'].strip())):
                        number_changes.append(change)
                    else:
                        text_changes.append(change)

                total_applied = 0
                total_failed = 0
                failed_changes = []
                
                # First pass: Apply pure number changes
                logging.info(f"Processing {len(number_changes)} number changes...")
                for change in number_changes:
                    try:
                        if self._apply_single_change(doc, change, retry=False):
                            total_applied += 1
                            doc.Save()  # Save after each successful change
                        else:
                            total_failed += 1
                            failed_changes.append(change)
                    except Exception as e:
                        logging.error(f"Error applying number change: {str(e)}")
                        total_failed += 1
                        failed_changes.append(change)
                
                # Second pass: Apply text changes
                logging.info(f"Processing {len(text_changes)} text changes...")
                for change in text_changes:
                    try:
                        if self._apply_single_change(doc, change, retry=False):
                            total_applied += 1
                            doc.Save()  # Save after each successful change
                        else:
                            total_failed += 1
                            failed_changes.append(change)
                    except Exception as e:
                        logging.error(f"Error applying text change: {str(e)}")
                        total_failed += 1
                        failed_changes.append(change)
                
                # Final retry for failed changes with relaxed thresholds
                if failed_changes:
                    logging.info(f"Retrying {len(failed_changes)} failed changes with relaxed matching...")
                    for change in failed_changes:
                        try:
                            if self._apply_single_change(doc, change, retry=True):
                                total_applied += 1
                                total_failed -= 1
                                doc.Save()
                        except Exception as e:
                            logging.error(f"Error in retry: {str(e)}")
                
                logging.info(f"Applied {total_applied} changes, {total_failed} failed")
                
                # Enable track changes for final save
                doc.TrackRevisions = True
                doc.Save()
                
                # Create timestamped backup of successful changes
                if total_applied > 0:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_path = test_dir / f"{Path(doc_path).stem}_track_changes_{timestamp}.docx"
                    doc.SaveAs(str(backup_path.resolve()))
                    logging.info(f"Saved backup with track changes: {backup_path}")
                
                return total_failed == 0 or total_applied > total_failed  # Success if most changes applied
                
            finally:
                doc.Close(SaveChanges=True)
                
        except Exception as e:
            logging.error(f"Error applying changes to document: {str(e)}")
            return False

    def _calculate_context_score(self, para, context):
        """Calculate context match score"""
        try:
            score = 0
            
            # Check previous paragraph
            if context['before']:
                prev_para = para.Previous
                if prev_para:
                    prev_score = self._calculate_text_similarity(
                        context['before'], 
                        {'text': prev_para.Range.Text.strip()}
                    )
                    score += prev_score * 0.5
                    
            # Check next paragraph
            if context['after']:
                next_para = para.Next
                if next_para:
                    next_score = self._calculate_text_similarity(
                        context['after'],
                        {'text': next_para.Range.Text.strip()}
                    )
                    score += next_score * 0.5
                    
            return score
            
        except Exception as e:
            logging.debug(f"Error calculating context score: {str(e)}")
            return 0

    def _calculate_structure_score(self, para, context):
        """Calculate structural similarity score"""
        try:
            score = 0
            struct = context.get('structure', {})
            
            # Compare style
            if struct.get('style') and hasattr(para, 'Style'):
                if para.Style.NameLocal == struct['style']:
                    score += 0.4
                    
            # Compare outline level
            if struct.get('level') is not None and hasattr(para, 'OutlineLevel'):
                if para.OutlineLevel == struct['level']:
                    score += 0.2
                    
            # Compare alignment
            if struct.get('alignment') is not None and hasattr(para, 'Alignment'):
                if para.Alignment == struct['alignment']:
                    score += 0.2
                    
            # Compare line spacing
            if struct.get('line_spacing') and hasattr(para, 'LineSpacing'):
                if abs(para.LineSpacing - struct['line_spacing']) < 0.1:
                    score += 0.2
                    
            return score
            
        except Exception as e:
            logging.debug(f"Error calculating structure score: {str(e)}")
            return 0

    def _calculate_position_score(self, para, context):
        """Calculate position-based similarity score"""
        try:
            # Get relative position in document
            doc = para.Range.Document
            total_paras = doc.Paragraphs.Count
            current_index = 0
            
            for i in range(total_paras):
                if doc.Paragraphs.Item(i + 1).Range.Start == para.Range.Start:
                    current_index = i
                    break
                    
            # Calculate relative position (0-1)
            relative_pos = current_index / total_paras if total_paras > 0 else 0
            
            # Score based on position - prefer similar relative positions
            # This helps maintain document structure between versions
            target_pos = context.get('relative_position', relative_pos)
            position_diff = abs(target_pos - relative_pos)
            
            return max(0, 1 - position_diff)
            
        except Exception as e:
            logging.debug(f"Error calculating position score: {str(e)}")
            return 0

def main():
    """Main entry point"""
    try:
        # Create log directory
        log_dir = Path('test')
        log_dir.mkdir(exist_ok=True)
        
        # Initialize logging with timestamp
        log_path = init_logging(log_dir / 'sync_changes.log')
        logging.info("Starting track changes synchronization")
        
        # Find matching document pairs
        pairs = find_matching_documents()
        if not pairs:
            logging.error("No document pairs found to process")
            return
            
        # Process each pair
        for english_doc, chinese_doc in pairs:
            logging.info("\nProcessing document pair:")
            logging.info(f"English document: {english_doc.name}")
            logging.info(f"Chinese document: {chinese_doc.name}")
            
            try:
                tracker = DocumentTracker()
                tracker.process_document_pair(str(english_doc), str(chinese_doc))
            except Exception as e:
                logging.error(f"Error processing pair: {str(e)}")
                continue
                
    except Exception as e:
        logging.error(f"Unhandled error: {str(e)}", exc_info=True)
        
if __name__ == "__main__":
    main()