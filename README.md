# sha3_vba
An Excel/VBA project to have SHA3 available in VBA. I need it as a building block to connect to the Ethereum blockchain, but it is not available through standard System.Security.Cryptography. A massive shout-out to [Chris Veness] https://github.com/chrisveness , as I basically took his Javascript example and "translated" it into VBA.

# Warning
This code is highly experimental and has hardly any error checks. There is probably a good reason why MS didn't include SHA3 into their normal Cryptography pack, although I didn't find it.

# How to use?
Import both .bas files you need. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you. The test code uses: https://github.com/vba-tools/vba-test

# ToDo
- Better error handling
- Clean up the code
- SHA3 with a password?
- Expand some bits I don't understand yet :)

# Done
- Standard VBA Not, Xor, And & Or are limited to Long numbers, made functions that also take bigger numbers
- Implemented bitwise shifting: shift left, shift right and shift right-zero-fill
- Decimal to Binary function, including unsigned numbers and numbers bigger than a Long
- Binary to Decimal function, including unsigned numbers and numbers bigger than a Long
- Implemented SHA3 - 224, 256, 384 and 512

# Donate
If this project saves you a lot of programming time, consider sending me a coffee or a beer:<br/>
ETH (or ERC-20 tokens): 0x9070C5D93ADb58B8cc0b281051710CB67a40C72B<br/>
<b>Cheers!</b>