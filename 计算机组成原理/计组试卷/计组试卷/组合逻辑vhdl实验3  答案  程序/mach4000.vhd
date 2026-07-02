library ieee;
use ieee.std_logic_1164.all;
use ieee.std_logic_arith.all;
use ieee.std_logic_unsigned.all;

--***********************Entity declaration*********************
entity Mach_entity is                       
  port(Clk,Reset:in std_logic;
       C,Z,S    :in std_logic;
       F15,C15  :in std_logic;
       Ir15,Ir14,Ir13,Ir12,Ir11,Ir10,Ir9,Ir8:in std_logic;
       Ir7,Ir6,Ir5,Ir4,Ir3,Ir2,Ir1,Ir0      :in std_logic;
       T3,T2,T1,T0:buffer std_logic;
       Mio,Req,Wr :Buffer std_logic;
       I8,I6,I5,I4,I3,I2,I1,I0:out std_logic;
       I7 :buffer std_logic;
       B,A:out    std_logic_vector(3 downto 0);
       Sst2,Sst1,Sst0,Ssh,Sci1,Sci0       :buffer std_logic;
       DC2_2,DC2_1,DC2_0,DC1_2,DC1_1,DC1_0:buffer std_logic;
       Ram15,Ram0,Q15,Q0,Cin:buffer std_logic;
       HndIns,Link,Bit8     :in std_logic;
       GswtoIbh,GswtoIbl,GIobush,GIobusl:out std_logic;
       GFtoIb,GAlutoIb,GIrltoIb         :out std_logic;
       MACH,C_M     :in    std_logic;
       Girh,Girl,GAR:inout std_logic;
       ARHCLK,ARLCLK:out   std_logic;  --garh and garl合并为GAR 用garh 77的脚控制，空出的garl改为ARHclk 80
       IRCLK        :inout std_logic);   --104,68,105,80  add 20071107
end Mach_entity;
--********************Architecture Declaration**********************
architecture Mach_archi of Mach_entity is
--******************A set instruction operating code****************
  CONSTANT ADD :STD_LOGIC_VECTOR(7 DOWNTO 0):="00000000";
  CONSTANT SUB :STD_LOGIC_VECTOR(7 DOWNTO 0):="00000001";
  CONSTANT AND1:STD_LOGIC_VECTOR(7 DOWNTO 0):="00000010";
  CONSTANT CMP :STD_LOGIC_VECTOR(7 DOWNTO 0):="00000011";
  CONSTANT XOR1:STD_LOGIC_VECTOR(7 DOWNTO 0):="00000100";
  CONSTANT TEST:STD_LOGIC_VECTOR(7 DOWNTO 0):="00000101";
  CONSTANT OR1 :STD_LOGIC_VECTOR(7 DOWNTO 0):="00000110";  
  CONSTANT MVRR:STD_LOGIC_VECTOR(7 DOWNTO 0):="00000111"; 
  CONSTANT DEC :STD_LOGIC_VECTOR(7 DOWNTO 0):="00001000";
  CONSTANT INC :STD_LOGIC_VECTOR(7 DOWNTO 0):="00001001";
  CONSTANT SHL :STD_LOGIC_VECTOR(7 DOWNTO 0):="00001010";
  CONSTANT SHR :STD_LOGIC_VECTOR(7 DOWNTO 0):="00001011";
  CONSTANT JR  :STD_LOGIC_VECTOR(7 DOWNTO 0):="01000001";
  CONSTANT JRC :STD_LOGIC_VECTOR(7 DOWNTO 0):="01000100";
  CONSTANT JRNC:STD_LOGIC_VECTOR(7 DOWNTO 0):="01000101";
  CONSTANT JRZ :STD_LOGIC_VECTOR(7 DOWNTO 0):="01000110";
  CONSTANT JRNZ:STD_LOGIC_VECTOR(7 DOWNTO 0):="01000111";
--******************B set instruction operating code****************
  CONSTANT JMPA:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000000";
  CONSTANT LDRR:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000001";
  CONSTANT IN1 :STD_LOGIC_VECTOR(7 DOWNTO 0):="10000010";
  CONSTANT STRR:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000011";
  CONSTANT PSHF:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000100";
  CONSTANT PUSH:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000101";
  CONSTANT OUT1:STD_LOGIC_VECTOR(7 DOWNTO 0):="10000110";
  CONSTANT POP :STD_LOGIC_VECTOR(7 DOWNTO 0):="10000111";
  CONSTANT MVRD:STD_LOGIC_VECTOR(7 DOWNTO 0):="10001000";
  CONSTANT POPF:STD_LOGIC_VECTOR(7 DOWNTO 0):="10001100";
  CONSTANT RET :STD_LOGIC_VECTOR(7 DOWNTO 0):="10001111";
--******************C set instruction operating code****************
  CONSTANT CALA:STD_LOGIC_VECTOR(7 DOWNTO 0):="11001110"; 
  signal T_temp:std_logic_vector(3 downto 0);
  signal GARH,GARL:STD_LOGIC;  --add 20071107
begin
--******************Clock Time generator***********************
  Timing_p:process(Clk,Reset)
    variable temp:std_logic_vector(1 downto 0);
  begin
    temp:=(Ir14,Ir11);
    if((reset='0')and(MACH='0')AND(C_M='1'))then T_temp<="1000";
    elsif((Clk 'event)and(Clk='1'))then 
         case T_temp is
           when "1000"=>T_temp<="0000";
           when "0000"=>T_temp<="0010";
           when "0010"=>if(Ir15='0')
                          then T_temp<="0011";
                          else T_temp<="0110";
                        end if;
           when "0110"=>case temp is
                          when "10"   =>T_temp<="0111";
                          when others =>T_temp<="0100";
                        end case;
           when "0100"=>if (Ir14='0')
                          then T_temp<="0000";
                          else T_temp<="0111";
                        end if;                        
           when "0111"=>T_temp<="0101";
           when others=>T_temp<="0000";
         end case;
     end if;
T3<=T_temp(3);T2<=T_temp(2);T1<=T_temp(1);T0<=T_temp(0); 
end process timing_p;
--******************Combined logic generator***********************
  Combin_p: process(T_temp)
    variable Comb:std_logic_vector(31 downto 0);
    variable instru:std_logic_vector(7 downto 0);
    variable Dr,Sr:std_logic_vector(3 downto 0);  
  begin
    instru:=Ir15&Ir14&Ir13&Ir12&Ir11&Ir10&Ir9&Ir8;
    Dr:=Ir7&Ir6&Ir5&Ir4;
    Sr:=Ir3&Ir2&Ir1&Ir0;
    Comb:=(others=>'0');
    case T_temp is
      when "1000"=>Comb :="10001100100101010101000001111000";
      when "0000"=>Comb :="10001000001101010101000001011000";
      when "0010"=>Comb :="00100100000000000000000000001000";
      when others=>Null;
    end case;  
    if (Ir13='0')and(C_M='1')then
      case T_temp is
        when "0011"=>case instru is     
                       when ADD =>Comb:="100011000001"&Dr&Sr&"001000000000";
                       when SUB =>Comb:="100011001001"&Dr&Sr&"001001000000";
                       when AND1=>Comb:="100011100001"&Dr&Sr&"001000000000";
                       when CMP =>Comb:="100001001001"&Dr&Sr&"001001000000";
                       when XOR1=>Comb:="100011110001"&Dr&Sr&"001000000000";
                       when TEST=>Comb:="100001100001"&Dr&Sr&"001000000000";
                       when OR1 =>Comb:="100011011001"&Dr&Sr&"001000000000";
                       when MVRR=>Comb:="100011000100"&Dr&Sr&"000000000000";
                       when DEC =>Comb:="100011001011"&Dr&"0000001000000000";
                       when INC =>Comb:="100011000011"&Dr&"0000001001000000";
                       when SHL =>Comb:="100111000011"&Dr&"0000110100000000";
                       when SHR =>Comb:="100101000011"&Dr&"0000101100000000";
                       when JR  =>Comb:="10001100010101010101000000000010";
                       when JRC =>Comb:="1000"&C&"100010101010101000000000010";
                       when JRNC=>Comb:="1000"&not C&"100010101010101000000000010";
                       when JRZ =>Comb:="1000"&Z&"100010101010101000000000010";
                       when JRNZ=>Comb:="1000"&not Z&"100010101010101000000000010"; 
                       when others=>NULL;
                     end case;
        when "0110"=>case instru is
                       when JMPA=>Comb:="10001000001101010101000001011000";
                       when LDRR=>Comb:="1000010001000000"&Sr&"000000011000";
                       when IN1 =>Comb:="10000100011100000000000000011010";
                       when STRR=>Comb:="100001000011"&Dr&"0000000000011000";
                       when PSHF=>Comb:="10001100101101000000000000011000";
                       when PUSH=>Comb:="10001100101101000000000000011000";
                       when OUT1=>Comb:="10000100011100000000000000011010";
                       when POP =>Comb:="10001000001101000100000001011000";
                       when MVRD=>Comb:="10001000001101010101000001011000";
                       when POPF=>Comb:="10001000001101000100000001011000";
                       when RET =>Comb:="10001000001101000100000001011000";
                       when CALA=>Comb:="10001000001101010101000001011000";
                       when OTHERS=>NULL; 
                     end case;
        when "0100"=>case instru is
                       when JMPA=>Comb:="00101100011101010000000000000000";
                       when LDRR=>Comb:="001011000111"&Dr&"0000000000000000";
                       when IN1 =>Comb:="01101100011100000000000000000000";
                       when STRR=>Comb:="0000010001000000"&Sr&"000000000001";
                       when PSHF=>Comb:="00000100000000000000000000000011";
                       when PUSH=>Comb:="0000010001000000"&Sr&"000000000001";
                       when OUT1=>Comb:="01000100010000000000000000000001";
                       when POP =>Comb:="001011000111"&Dr&"0000000000000000";
                       when MVRD=>Comb:="001011000111"&Dr&"0000000000000000";
                       when POPF=>Comb:="00100100000000000000010000000000";
                       when RET =>Comb:="00101100011101010000000000000000";
                       when CALA=>Comb:="00100000011100000000000000000000";
                       when others=>NULL;
                     end case;
        when "0111"=>if instru=CALA then
                     Comb:="10001100101101000000000000011000";
                   end if;
        when "0101"=>if instru=CALA then
                     Comb:="00001000001001010101000000000001";
                   end if;
        when others=>NULL;
      end case; 
    else NULL;--Extended Instruction
    end if;  
--**************Combined logic signal assignment******************
Mio<=Comb(31);Req<=Comb(30);Wr<=Comb(29);
I8 <=Comb(28);I7 <=Comb(27);I6<=Comb(26);
I5 <=Comb(25);I4 <=Comb(24);I3<=Comb(23);
I2 <=Comb(22);I1 <=Comb(21);I0<=Comb(20);
A  <=Comb(15 downto 12);
B  <=Comb(19 downto 16);
Sst2<=Comb(11);Sst1<=Comb(10);Sst0<=Comb(9);
Ssh <=Comb(8) ;Sci1<=Comb(7) ;Sci0<=Comb(6);
DC2_2<=Comb(5);DC2_1<=Comb(4);DC2_0<=Comb(3);
DC1_2<=Comb(2);DC1_1<=Comb(1);DC1_0<=Comb(0);
end process Combin_p;
--******************DC1\DC2 encode***********************
ctr_b:Block
  begin
      Gswtoibh<=(not HndIns or Mio or not Wr) and Link;
      Gswtoibl<=(not HndIns or Mio or not Wr) and Link;
      Giobush   <=((not Link) or HndIns or Mio or MACH);
      Giobusl   <=((not Link) or HndIns or Mio or MACH);

      ARHCLK<=(not GARH)and IRCLK and(not Bit8);
      ARLCLK<=(not GARH)and IRCLK and(not Bit8);  
      GAR   <=MACH;
      
      GirltoIb<=     (Dc1_0 or(not Dc1_1)or Dc1_2);
      GalutoIb<=((not(Dc1_0)or     Dc1_1 or Dc1_2));
      GftoIb  <=( not(Dc1_0)or not(Dc1_1)or Dc1_2);
      
      Girh<=(not(Dc2_0) or Dc2_1 or Dc2_2) or Bit8;
      Girl<=(not(Dc2_0)or Dc2_1 or Dc2_2);

      GARL<=(not(Dc2_0) or not(Dc2_1) or Dc2_2);
      GARH<=(   (Dc2_0  or not(Dc2_1) or Dc2_2)or not(Bit8))
            and (((not Dc2_0) or (not Dc2_1) or Dc2_2)or Bit8);
end block ctr_b;
--******************Shift signal***********************
shift:block
  begin  
    Cin  <=((not Ssh)and(not Sci1)and Sci0) or
           ((not Ssh)and Sci1 and (not Sci0)and C);
           
    Q0   <=(Ssh and Sci1 and (not Sci0) and (not F15)) when I7='1' else 'Z';
    Q15  <=((Ssh and Sci1 and (not Sci0) and Ram0) or
          (Ssh and Sci1 and Sci0 and Ram0)) when I7='0' else 'Z';
       
    Ram0 <=(Ssh and (not Sci1) and Sci0 and C) or 
           (Ssh and Sci1 and (not Sci0) and Q15)  when I7='1' else 'Z';
    Ram15<=((Ssh and (not Sci1) and Sci0 and C) or
           (Ssh and Sci1 and (not Sci0) and C15) or
           (Ssh and Sci1 and Sci0 and F15)) when I7='0' else 'Z';
  end block shift;
end Mach_archi;

