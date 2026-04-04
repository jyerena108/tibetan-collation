from Pydurma.gen.normalizer_gen import GenericNormalizer
from Pydurma.gen.tokenizer_gen import GenericTokenizer
from Pydurma.encoder import Encoder
from Pydurma.aligners.fdmp import FDMPaligner
from Pydurma.utils.utils import column_matrix_to_row_matrix, token_row_to_text_row


def collate_texts(witness_texts):
    """
    witness_texts: dict like
      {
         "W1": "tibetan text version 1",
         "W2": "tibetan text version 2",
         ...
      }
    """
    normalizer = GenericNormalizer()
    encoder = Encoder()
    tokenizer = GenericTokenizer(encoder, normalizer)
    aligner = FDMPaligner()

    token_lists = {}
    token_strings = {}

    # normalize + tokenize each witness
    for wname, text in witness_texts.items():
        tokens, tokenstr = tokenizer.tokenize(text)
        token_lists[wname] = tokens
        token_strings[wname] = tokenstr

    # align
    matrix = aligner.get_alignment_matrix(token_strings, token_lists)

    # convert to row matrix
    row_matrix = column_matrix_to_row_matrix(matrix)

    # rebuild human-readable rows for each witness
    readable_rows = {}
    for i, (wname, text) in enumerate(witness_texts.items()):
        readable_rows[wname] = token_row_to_text_row(row_matrix[i], text)

    return readable_rows


if __name__ == "__main__":
    # tiny Tibetan test – replace with real lines when you’re ready
    witnesses = {
        "W1": "བཀྲ་ཤིས་བདེ་ལེགས།",
        "W2": "བཀྲ་ཤིས་བདེ་ལེགས",
        "W3": "བཀྲ་ཤིས་བདེ་ལེགས་ཀྱིས།",
    }

    result = collate_texts(witnesses)

    print("\n=== Collation Result ===")
    for wname, row in result.items():
        print(f"{wname}: {row}")
