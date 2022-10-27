#!/usr/bin/Rscript --no-init-file
# vim: fileencoding=utf-8

# Project:  spreadsheet-hash
# Version:  1.0
# Created:  26-10-2022
# Authors:  Leonardo Gama
# Homepage: github.com/leogama/spreadsheet-hash

# MIT License
#
# Copyright  2022  Leonardo dos Reis Gama
#
# Permission is hereby granted, free of charge, to any person obtaining a
# copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.

library(bitops)

FNV32_BASIS <- 0x811C9DC5
FNV32_PRIME_HI <- 0x01000000
FNV32_PRIME_LO <- 0x0193
MASK_16_BITS <- 0xFFFF
MASK_32_BITS <- 2**32  # modulo masking

fnv1a_32 <- function(text, length = 8L, salt = NULL) {
    length <- as.integer(length)
    if (is.na(length) || length < 1L || length > 8L )
        stop("'length' must be an integer between 1 and 8 (inclusive)")
    if (length(text) != 1L)
        stop("'text' must have length 1")

    if (!is.null(salt))
        text <- paste0(salt, text)
    text <- iconv(text, to = 'UTF-8', toRaw = TRUE)[[1L]]

    hash <- FNV32_BASIS
    for (byte in as.numeric(text)) {
        hash <- hash %^% byte
        hash_lo <- hash %&% MASK_16_BITS  # avoid overflow
        hash <- hash * FNV32_PRIME_LO + hash_lo * FNV32_PRIME_HI
        hash <- hash %% MASK_32_BITS
    }

    if (length < 8L) {
        # XOR-folding
        bit_size <- length * 4L
        size_mask <- (1 %<<% bit_size) - 1
        hash <- ((hash %>>% bit_size) %^% hash) %&% size_mask
        fmt <- paste0("%0", length, "X")
        return(sprintf(fmt, hash))
    } else {
        # 0xFFFFFFFF can't be represented by integer type
        hash_hi <- hash %>>% 16L
        hash_lo <- hash %&% MASK_16_BITS
        return(sprintf("%04X%04X", hash_hi, hash_lo))
    }
}

name2id <- function(name, length = 8L, salt = NULL) {
    name <- trimws(toupper(name))
    name <- gsub('\\.', ' ', name)
    name <- gsub('\\s+', ' ', name)
    name <- iconv(name, to = 'ASCII//TRANSLIT')
    fnv1a_32(name, length)
}

if (!interactive() && endsWith(tail(commandArgs(), 1L), 'name2id.R')) {
    hash1 = fnv1a_32("hello world")
    hash2 = fnv1a_32("world", salt = "hello ")
    cat(sprintf("%s == %s\n", hash1, hash2))
    hash1 = name2id("Maria  D.Assunção ", 4)
    hash2 = name2id("MARIA D ASSUNCAO", 4)
    cat(sprintf("    %s == %s\n", hash1, hash2))
}
