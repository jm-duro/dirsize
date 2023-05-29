-- Creative Commons Licence
-- Luca De Feo 2016, licensed under CC-BY-SA 4.0.
-- http:--defeo.lu/in420/Coder%20et%20d%C3%A9coder%20en%20UTF-8%20et%20UTF-16

-- Constantes liees aux plages UTF-16
constant
  LEAD_START = #D800, 
  TAIL_START = #DC00, 
  TAIL_END   = #DFFF,
  BMP_END    = #FFFF,
  MAX_CP     = #10FFFF
        
-- Constantes liees aux plages UTF-8
constant
  UTF8_CP1 = #80,
  UTF8_CP2 = #800,
  UTF8_CP3 = #10000,
  UTF8_CP4 = #200000

constant
  UTF8_BX = #80,
  UTF8_B2 = #C0,
  UTF8_B3 = #E0,
  UTF8_B4 = #F0,
  UTF8_B5 = #F8

------------------------------------------------------------------------------

/*
  Convertit deux mots de 16 bits (supposes encodes en UTF-16) en
  un codepoint Unicode. Le deuxieme mot est ignore le cas echeant.

  Renvoie -1 s'il y a un probleme d'encodage.
*/
function cp_from_UTF16(integer lead, integer tail)
  if (lead < LEAD_START) or (lead > TAIL_END) then
    return lead
  elsif (lead >= LEAD_START and lead < TAIL_START)
    return (lead - #D800) * #400 + (tail - #DC00) + #10000
  else 
    return -1
  end if
end function

------------------------------------------------------------------------------

/*
  Convertit un codepoint Unicode en son encodage UTF-16 (un
  tableau de un ou deux entiers de 16 bits).

  Renvoie null si le codepoint est invalide.
*/
function cp_to_UTF16(integer codepoint)
  if (codepoint >= 0) and (codepoint < BMP_END) then
    return { codepoint }
  elsif (codepoint > BMP_END) and (codepoint <= MAX_CP) then
    return { (codepoint - #10000) / #400 + #D800,
             remainder((codepoint - #10000), #400) + 0XDC00 }
  else
    return 0
  end if
end function

------------------------------------------------------------------------------

/*
  Convertit un codepoint Unicode en son encodage UTF-8 (un
  tableau de un a quatre entiers de 8 bits).

  Renvoie null si le codepoint est invalide.
*/
function cp_to_UTF8(integer codepoint)
  if (codepoint >= 0) and (codepoint < UTF8_CP1) then
    return { codepoint }
  elsif (codepoint < UTF8_CP2) then
    return { codepoint / #40 + UTF8_B2,
             remainder(codepoint, #40) + UTF8_BX }
  elsif (codepoint < UTF8_CP3) then
    return { codepoint / #40 / #40 + UTF8_B3,
             remainder(codepoint / #40, #40) + UTF8_BX,
             remainder(codepoint, #40) + UTF8_BX }
      return bytes;
  elsif (codepoint < UTF8_CP4) then
      return { codepoint / #40 / #40 / #40 + UTF8_B4,
               remainder(codepoint / #40 / #40, #40) + UTF8_BX,
               remainder(codepoint / #40, #40) + UTF8_BX,
               remainder(codepoint, #40) + UTF8_BX }
  else
    return 0
  end if
end function

------------------------------------------------------------------------------

/*
  Convertit quatre mots de 8 bits (supposes encodes en UTF-8) en
  un codepoint Unicode. Les mots en plus sont ignores le cas
  echeant.

  Renvoie -1 s'il y a un probleme d'encodage.
*/
function cp_from_UTF8(integer b1, integer b2, integer b3, integer b4)
  if (b1 < UTF8_BX) then
    return b1
  elsif (b1 < UTF8_B2)
    return -1
  elsif (b1 < UTF8_B3) and (b2 >= UTF8_BX) and (b2 < UTF8_B2) then
    return remainder(b1, #20)*#40 + remainder(b2, #40)
  elsif (b1 < UTF8_B4) and (b2 >= UTF8_BX) and (b2 < UTF8_B2) and
        (b3 >= UTF8_BX) and (b3 < UTF8_B2) then
    return remainder(b1, #10)*#40*#40 + remainder(b2, #40)*#40 +
           remainder(b3, #40)
  elsif (b1 < UTF8_B5) and (b2 >= UTF8_BX) and (b2 < UTF8_B2) and
        (b3 >= UTF8_BX) and (b3 < UTF8_B2) and (b4 >= UTF8_BX) and
        (b4 < UTF8_B2) then
    return remainder(b1, #8)*#40*#40*#40 + remainder(b2, #40)*#40*#40 +
           remainder(b3, #40)*#40 + remainder(b4, #40)
  else
    return -1
end function

------------------------------------------------------------------------------

/*
  Convertit un flux de UTF-8 a UTF-16
*/
public static void utf8to16(InputStream in, OutputStream out)
    throws IOException {

    integer b1 = in.read(), 
        b2 = in.read(), 
        b3 = in.read(),
        b4;
    while (b1 != -1) {
        b4 = in.read();
        integer cp = cp_from_UTF8(b1, b2, b3, b4);
        if (cp != -1) {
    	integer[] pairs = cp_to_UTF16(cp);
    	for (integer i = 0 ; i < pairs.length ; i++) {
    	    out.write(pairs[i] / #100);
    	    out.write(pairs[i] % #100);
    	}
        }
        b1 = b2; b2 = b3; b3 = b4;
    }
}

------------------------------------------------------------------------------

/*
  Convertit un flux de UTF-16 a UTF-18
*/
public static void utf16to8(InputStream in, OutputStream out) 
    throws IOException {
    integer b1 = in.read(), 
        b2 = in.read(), 
        b3, b4;

    while (b1 != -1) {
        b3 = in.read();
        b4 = in.read();
        integer cp = cp_from_UTF16(b1 * #100 + b2, b3 * #100 + b4);
        if (cp != -1) {
    	integer[] bytes = cp_to_UTF8(cp);
    	for (integer i = 0 ; i < bytes.length ; i++)
    	    out.write(bytes[i]);
        }
        b1 = b3; b2 = b4;
    }
}

public static void main(String[] args) 
    throws IOException, UnsupportedEncodingException {
    InputStream in = System.in;
    OutputStream out = System.out;
    integer utf16 = 0;

    if (args.length >= 1 and args[0].equals("-utf16"))
        utf16 = 1;
    if (args.length >= 1 + utf16)
        in = new FileInputStream(args[utf16]);
    if (args.length >= 2 + utf16) 
        out = new FileOutputStream(args[1+utf16]);

    if (utf16 == 0)
        utf8to16(in, out);
    else
        utf16to8(in, out);
}

