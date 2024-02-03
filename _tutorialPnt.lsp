; //@#$!@#$@!#$Progress### << ///Jan_2023..??>> @ @#$@%#!@%!Saving??@!#$@!#$@!#$!@#$@!#$
; //!!!!! don't stay at one part for more than 1 weekds...!!!!
; //!!!!! proceed as you go through by commenting .. for furhter refirinement..
; //..//...//..//...//..//..//..//...//paly around with it thats how you get it
; //..//..Code explanation.//..use...//.ChatGPT.//..//..//...//
; //..//...//..//...//..//..//..//...//

#| |#


#| 
;  Basic Syntax 


	(write-line "Hello World") ; greet the world

	; tell them your whereabouts

	(write-line "I am at 'Tutorials Point'! Learning LISP")

;Numbers

	;Naming Conventions in LISP

	;Create a file named main.lisp and type the following code into it.

	(write-line "single quote used, it inhibits evaluation")
	(write '(* 2 3))
	(write-line " ")
	(write-line "single quote not used, so expression evaluated")
	(write (* 2 3))


	;single quote used, it inhibits evaluation
	(* 2 3) 

|#



#| 

;Data types

;Example 1
;Create new source code file named main.lisp and type the following code in it.

(setq x 10)
(setq y 34.567)
(setq ch nil)
(setq n 123.78)
(setq bg 11.0e+4)
(setq r 124/2)

(print x)
(print y)
(print n)
(print ch)
(print bg)
(print r)



;Example 2
;Next let's check the types of the variables used in the previous example. Create new source code file named main. lisp and type the following code in it.

(defvar x 10)
(defvar y 34.567)
(defvar ch nil)
(defvar n 123.78)
(defvar bg 11.0e+4)
(defvar r 124/2)

(print (type-of x))
(print (type-of y))
(print (type-of n))
(print (type-of ch))
(print (type-of bg))
(print (type-of r))


|#


#| |#
;Macros

;Defining a Macro
;(defmacro macro-name (parameter-list))

; "Optional documentation string."

; body-form

; The macro definition consists of the name of the macro, 
	; parameter list, 
	; an optional documentation string, and 
	; a body of Lisp expressions that defines the job to be performed by the macro.

;Example

(defmacro setTo10(num)
(setq num 10)(print num))
(setq x 25)
(print x)
(setTo10 x)

#| |#
;Variables


;Global Variables with permanent values 

; declared using the defvar construct.

;For example
(defvar x 234)
(write x)

;Since there is no type declaration for variables in LISP, you directly specify a value for a symbol with the setq construct.

;For Example
(setq x 10)

;The above expression assigns the value 10 to the variable x. 

;The symbol-value function allows you to extract the value stored at the symbol storage place.

;For Example
;Create new source code file named main.lisp and type the following code in it.

(setq x 10)
(setq y 20)
(format t "x = ~2d y = ~2d ~%" x y)

(setq x 100)
(setq y 200)
(format t "x = ~2d y = ~2d" x y)

#| |#
;Local Variables
;	Local variables are defined within a given procedure. 


;let and prog for creating local variables.

(let ((var1  val1) (var2  val2).. (varn  valn)))



;Example
;Create new source code file named main.lisp and type the following code in it.

(let ((x 'a) (y 'b)(z 'c))
(format t "x = ~a y = ~a z = ~a" x y z))

;
x = A y = B z = C

;
;Example


(prog ((x '(a b c))(y '(1 2 3))(z '(p q 10)))
(format t "x = ~a y = ~a z = ~a" x y z))


x = (A B C) y = (1 2 3) z = (P Q 10)




#| |#
;Constants
; constants are variables that never change 

;Example
;


(defconstant PI 3.141592)
(defun area-circle(rad)
  (terpri)
  (format t "Radius: ~5f" rad)
  (format t "~%Area: ~10f" (* PI rad rad)))
(area-circle 10)


#| |#
; Operators


;
;The operations allowed on data could be categorized as:
;
;Arithmetic Operations
;Comparison Operations
;Logical Operations
;Bitwise Operations
;Arithmetic Operations

;try fefniing A and B here...

;Operator	Description	Example

(+A B) 	;	
(- A B) ;	
(* A B) ;	
(/ A B) ;	

(decf A 4) ;	Modulus Operator and 
(incf A 3) ;	Increments operator 
(mod B A ) ;	Decrements operator 

#| |#
;Comparison Operations



;Assume variable A holds 10 and variable B holds 20, then:

;Show Examples

;Operator	Description	Example

(/= A B) ; is not equal.
(= A B)  ;/=	is equal 
(< A B)  ;>	
(> A B)  ;<	
(>= A B) ;>=
(<= A B) ;<=

(max A B) ; returns 20
(min A B) ; returns 20

#| |#
;Logical Operations on Boolean Values

;Bitwise Operations on Numbers
;Bitwise operators work on bits and perform bit-by-bit operation. The truth tables for bitwise and, or, and xor operations are as follows:
;
;Show Examples
;
;p	q	p and q	p or q	p xor q
;0	0	0	0	0
;0	1	0	1	1
;1	1	1	1	0
;1	0	0	1	1
;Assume if A = 60; and B = 13; now in binary format they will be as follows:
;A = 0011 1100
;B = 0000 1101
;-----------------
A and B = 0000 1100
A or B = 0011 1101
A xor B = 0011 0001
not A  = 1100 0011

;The Bitwise operators supported by LISP are listed in the following table. Assume variable A holds 60 and variable B holds 13, then:
;
;Description	Example
(lognor a b);logand	This returns the bit-wise logical AND of its arguments. If no argument is given, then the result is -1, which is an identity for this operation.			(logand a b)) will give 12
(logxor a b);logior	This returns the bit-wise logical INCLUSIVE OR of its arguments. If no argument is given, then the result is zero, which is an identity for this operation.	(logior a b) will give 61
(logior a b);logxor	This returns the bit-wise logical EXCLUSIVE OR of its arguments. If no argument is given, then the result is zero, which is an identity for this operation.	(logxor a b) will give 49
(logand a b);lognor	This returns the bit-wise NOT of its arguments. If no argument is given, then the result is -1, which is an identity for this operation.					(lognor a b) will give -62,


#| |#
;(#|Decisions
;
;)
;
;
;

#| |#
;Loops

;Gracefully Exiting From a Block

;The block and return-from allows you to exit gracefully from any nested blocks in case of any error.
;
;The block function 

; Syntax is:
;
;(block block-name(...))

;The return-from function takes a block name and an optional (the default is nil) return value.


;Example
;Create a new source code file named main.lisp and type the following code in it:
;

(defun demo-function (flag)
  (print 'entering-outer-block)
  
  (block outer-block
     (print 'entering-inner-block)
     (print (block inner-block
     
        (if flag
           (return-from outer-block 3)
           (return-from inner-block 5)
        )
        
        (print 'This-wil--not-be-printed))
     )
     
     (print 'left-inner-block)
     (print 'leaving-outer-block)
  t)
)
(demo-function t)
(terpri)
(demo-function nil)


#| |#
;(#|Functions
;ing Functions in LISP
;The macro named defun is used for defining functions. The defun macro needs three arguments:
;
;

(defun name (parameter-list) "Optional documentation string." body)

;Let us illustrate the concept with simple examples.
;

;Example 1
;Let's write a function named averagenum that will print the average of four numbers. We will send these numbers as parameters.
;
;Create a new source code file named main.lisp and type the following code in it.
;

(defun averagenum (n1 n2 n3 n4)
  (/ ( + n1 n2 n3 n4) 4)
)
(write(averagenum 10 20 30 40))


;Example 2
;Let's define and call a function that would calculate the area of a circle when the radius of the circle is given as an argument.
;
;Create a new source code file named main.lisp and type the following code in it.
;

(defun area-circle(rad)
  "Calculates area of a circle with given radius"
  (terpri)
  (format t "Radius: ~5f" rad)
  (format t "~%Area: ~10f" (* 3.141592 rad rad))
)
(area-circle 10)


;You can provide an empty list as parameters, which means the function takes no arguments, the list is empty, written as ().

#| |#
;Lambda Functions

;
;
;

#| |#
;|Predicates

;Ceate a new source code file named main.lisp and type the following code in it.
;

(write (atom 'abcd))
(terpri)
(write (equal 'a 'b))
(terpri)
(write (evenp 10))
(terpri)
(write (evenp 7 ))
(terpri)
(write (oddp 7 ))
(terpri)
(write (zerop 0.0000000001))
(terpri)
(write (eq 3 3.0 ))
(terpri)
(write (equal 3 3.0 ))
(terpri)
(write (null nil ))


;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;

(defun factorial (num)
  (cond ((zerop num) 1)
     (t ( * num (factorial (- num 1))))
  )
)
(setq n 6)
(format t "~% Factorial ~d is: ~d" n (factorial n))

#| |#
;Numbers

;arious Numeric Types in LISP
;The following table describes various number type data available in LISP:
;
;Data type	Description
;fixnum	
;bignum	
;ratio	
;float	
;complex

;Example
;Create a new source code file named main.lisp and type the following code in it.

(write (/ 1 2))
(terpri)
(write ( + (/ 1 2) (/ 3 4)))
(terpri)
(write ( + #c( 1 2) #c( 3 -4)))

#| |#
;Number Functions

;The following table describes some commonly used numeric functions:
;
;Function	Description

;+, -, *, /			
;sin, cos, tan, acos, asin, atan		
;sinh, cosh, tanh, acosh, asinh, atanh	
;exp	
;expt	
;sqrt	
;log	
;conjugate	It calculates the complex conjugate of a number. In case of a real number, it returns the number itself.
;abs	
;gcd	
;lcm	
;isqrt	
;floor, ceiling, truncate, round	
;ffloor, fceiling, ftruncate, fround	
;mod, rem	
;float	
;rational, rationalize	
;numerator, denominator	
;realpart, imagpart	

;Example
;Create a new source code file named main.lisp and type the following code in it.


(write (/ 45 78))
(terpri)
(write (floor 45 78))
(terpri)
(write (/ 3456 75))
(terpri)
(write (floor 3456 75))
(terpri)
(write (ceiling 3456 75))
(terpri)
(write (truncate 3456 75))
(terpri)
(write (round 3456 75))
(terpri)
(write (ffloor 3456 75))
(terpri)
(write (fceiling 3456 75))
(terpri)
(write (ftruncate 3456 75))
(terpri)
(write (fround 3456 75))
(terpri)
(write (mod 3456 75))
(terpri)
(setq c (complex 6 7))
(write c)
(terpri)
(write (complex 5 -9))
(terpri)
(write (realpart c))
(terpri)
(write (imagpart c))

#| |#
; Characters


;Example
;Create a new source code file named main.lisp and type the following code in it.
;
(write 'a)
(terpri)
(write #\a)
(terpri)
(write-char #\a)
(terpri)
(write-char 'a)



;Special Characters

;Common LISP allows using the following special characters in your code. They are called the semi-standard characters.
;
;#\Backspace
;#\Tab
;#\Linefeed
;#\Page
;#\Return
;#\Rubout

;Character Comparison Functions



;Case Sensitive Functions	Case-insensitive Functions	Description
;char=	char-equal		
;char/=	char-not-equal	
;char&#60;	char-lessp	
;char>	char-greaterp	
;char&#60;=	char-not-greaterp	
;char>=	char-not-lessp			

;Example
;Create a new source code file named main.lisp and type the following code in it.
;
; case-sensitive comparison

(write (char= #\a #\b))
(terpri)
(write (char= #\a #\a))
(terpri)
(write (char= #\a #\A))
(terpri)
  
;case-insensitive comparision
(write (char-equal #\a #\A))
(terpri)
(write (char-equal #\a #\b))
(terpri)
(write (char-lessp #\a #\b #\c))
(terpri)
(write (char-greaterp #\a #\b #\c))

#| |#
; Arrays


;For example, to access the content of the tenth cell, we write:
;
(aref my-array 9)

;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;

(write (setf my-array (make-array '(10))))
(terpri)
(setf (aref my-array 0) 25)
(setf (aref my-array 1) 23)
(setf (aref my-array 2) 45)
(setf (aref my-array 3) 10)
(setf (aref my-array 4) 20)
(setf (aref my-array 5) 17)
(setf (aref my-array 6) 25)
(setf (aref my-array 7) 19)
(setf (aref my-array 8) 67)
(setf (aref my-array 9) 30)
(write my-array)


;Example 2
;Let us create a 3-by-3 array.
;
;Create a new source code file named main.lisp and type the following code in it.
;

(setf x (make-array '(3 3) 
  :initial-contents '((0 1 2 ) (3 4 5) (6 7 8)))
)
(write x)

;Example 3
;Create a new source code file named main.lisp and type the following code in it.
;

(setq a (make-array '(4 3)))
(dotimes (i 4)
  (dotimes (j 3)
     (setf (aref a i j) (list i 'x j '= (* i j)))
  )
)
(dotimes (i 4)
  (dotimes (j 3)
     (print (aref a i j))
  )
)


;Complete Syntax for the make-array Function

;Argument	Description


;Example 4
;Create a new source code file named main.lisp and type the following code in it.
;

(setq myarray (make-array '(3 2 3) 
  :initial-contents 
  '(((a b c) (1 2 3)) 
     ((d e f) (4 5 6)) 
     ((g h i) (7 8 9)) 
  ))
) 
(setq array2 (make-array 4 :displaced-to myarray :displaced-index-offset 2)) 
(write myarray)
(terpri)
(write array2)


(setq myarray (make-array '(3 2 3) 
  :initial-contents 
  '(((a b c) (1 2 3)) 
     ((d e f) (4 5 6)) 
     ((g h i) (7 8 9)) 
  ))
) 
(setq array2 (make-array '(3 2) :displaced-to myarray :displaced-index-offset 2)) 
(write myarray)
(terpri)
(write array2)


(setq myarray (make-array '(3 2 3) 
  :initial-contents 
  '(((a b c) (1 2 3)) 
     ((d e f) (4 5 6)) 
     ((g h i) (7 8 9)) 
  ))
) 
(setq array2 (make-array '(3 2) :displaced-to myarray :displaced-index-offset 5)) 
(write myarray)
(terpri)
(write array2)



;Example 5
;Create a new source code file named main.lisp and type the following code in it.
;
;;a one dimensional array with 5 elements, 
;;initail value 5

(write (make-array 5 :initial-element 5))
(terpri)

;two dimensional array, with initial element a

(write (make-array '(2 3) :initial-element 'a))
(terpri)


;an array of capacity 14, but fill pointer 5, is 5

(write(length (make-array 14 :fill-pointer 5)))
(terpri)

;however its length is 14

(write (array-dimensions (make-array 14 :fill-pointer 5)))
(terpri)


; a bit array with all initial elements set to 1


(write(make-array 10 :element-type 'bit :initial-element 1))
(terpri)


; a character array with all initial elements set to a
; is a string actually

(write(make-array 10 :element-type 'character :initial-element #\a)) 
(terpri)


; a two dimensional array with initial values a

(setq myarray (make-array '(2 2) :initial-element 'a :adjustable t))
(write myarray)
(terpri)


;readjusting the array

(adjust-array myarray '(1 3) :initial-element 'b) 
(write myarray)

#| |#
; Strings


;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write-line "Hello World")
(write-line "Welcome to Tutorials Point")
;escaping the double quote character
(write-line "Welcome to \"Tutorials Point\"")


;Numeric comparison functions and operators, 
;
;The following table provides the functions:
;
;Case Sensitive Functions	Case-insensitive Functions	Description
;= 		string-equal		
;/=		string-not-equal	
;&#60;	string-lessp		
;>		string-greaterp		
;&#60;=	string-not-greaterp	
;>=		string-not-lessp	

;Example
;Create a new source code file named main.lisp and type the following code in it.
;
;; case-sensitive comparison

(write (string= "this is test" "This is test"))
(terpri)
(write (string> "this is test" "This is test"))
(terpri)
(write (string< "this is test" "This is test"))
(terpri)


;case-insensitive comparision

(write (string-equal "this is test" "This is test"))
(terpri)
(write (string-greaterp "this is test" "This is test"))
(terpri)
(write (string-lessp "this is test" "This is test"))
(terpri)


;checking non-equal

(write (string/= "this is test" "this is Test"))
(terpri)
(write (string-not-equal "this is test" "This is test"))
(terpri)
(write (string/= "lisp" "lisping"))
(terpri)
(write (string/= "decent" "decency"))

#| |#
;Case Controlling Functions

;The following table describes the case controlling functions:
;
;Function	Description
;string-upcase		Converts the string to upper case
;string-downcase	Converts the string to lower case
;string-capitalize	Capitalizes each word in the string

;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write-line (string-upcase "a big hello from tutorials point"))
(write-line (string-capitalize "a big hello from tutorials point"))


;Function	Description
					
;string-trim		
;String-left-trim	
;String-right-trim	
					
;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write-line (string-trim " " "   a big hello from tutorials point   "))
(write-line (string-left-trim " " "   a big hello from tutorials point   "))
(write-line (string-right-trim " " "   a big hello from tutorials point   "))
(write-line (string-trim " a" "   a big hello from tutorials point   "))


;Calculating Length
	;The length function 

;Extracting Sub-string
	;The subseq function 

;Accessing a Character in a String
	;The char function 


;Example
;
;Create a new source code file named main.lisp and type the following code in it.
;

(write (length "Hello World"))
(terpri)
(write-line (subseq "Hello World" 6))
(write (char "Hello World" 6))

#| |#
;Sorting and Merging of Strings

;Example
;
;Create a new source code file named main.lisp and type the following code in it.
;

;;sorting the strings

(write (sort (vector "Amal" "Akbar" "Anthony") #'string<))
(terpri)

;
;;merging the strings

(write (merge 'vector (vector "Rishi" "Zara" "Priyanka") (vector "Anju" "Anuj" "Avni") #'string<))


;Reversing a String

(write-line (reverse "Are we not drawn onward, we few, drawn onward to new era"))	; trying in python def(d):print(d[::-1])


;Concatenating Strings


;For example, Create a new source code file named main.lisp and type the following code in it.
;

(write-line (concatenate 'string "Are we not drawn onward, " "we few, drawn onward to new era"))

#| |#
;Sequences

;Creating a Sequence


;For example, Create a new source code file named main.lisp and type the following code in it.
;

(write (make-sequence '(vector float) 
  10 
  :initial-element 1.0))


;Generic Functions on Sequences

;Function	Description

;elt		
;length		
;subseq		
;copy-seq	
;fill	
;replace	
;count		
;reverse	
;nreverse	
;concatenate
;position	
;find	
;sort	
;merge	
;map	
;some	
;every	
;notany	
;notevery	
;reduce	
;search	
;remove	
;delete	
;substitute	
;nsubstitute	
;mismatch		

;Standard Sequence Function Keyword Arguments

;Argument	
;EQL 		
;NIL:		
;0			
;NIL		



;Finding Length and Element

;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(setq x (vector 'a 'b 'c 'd 'e))
(write (length x))
(terpri)
(write (elt x 3))


#| |#
;Modifying Sequences


;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;

(write (count 7 '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (remove 5 '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (delete 5 '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (substitute 10 7 '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (find 7 '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (position 5 '(1 5 6 7 8 9 2 7 3 4 5)))


;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;

(write (delete-if #'oddp '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (delete-if #'evenp '(1 5 6 7 8 9 2 7 3 4 5)))
(terpri)
(write (remove-if #'evenp '(1 5 6 7 8 9 2 7 3 4 5) :count 1 :from-end t))
(terpri)
(setq x (vector 'a 'b 'c 'd 'e 'f 'g))
(fill x 'p :start 1 :end 4)
(write x)

#| |#
;Sorting and Merging Sequences


;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;

(write (sort '(2 4 7 3 9 1 5 4 6 3 8) #'<))
(terpri)
(write (sort '(2 4 7 3 9 1 5 4 6 3 8) #'>))
(terpri)


;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;

(write (merge 'vector #(1 3 5) #(2 4 6) #'<))
(terpri)
(write (merge 'list #(1 3 5) #(2 4 6) #'<))
(terpri)

#| |#
;Sequence Predicates

;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write (every #'evenp #(2 4 6 8 10)))
(terpri)
(write (some #'evenp #(2 4 6 8 10 13 14)))
(terpri)
(write (every #'evenp #(2 4 6 8 10 13 14)))
(terpri)
(write (notany #'evenp #(2 4 6 8 10)))
(terpri)
(write (notevery #'evenp #(2 4 6 8 10 13 14)))
(terpri)

#| |#
;Mapping Sequences


;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write (map 'vector #'* #(2 3 4 5) #(3 5 4 8)))

#| |#
;Lists


;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(write (cons 1 2))
(terpri)
(write (cons 'a 'b))
(terpri)
(write (cons 1 nil))
(terpri)
(write (cons 1 (cons 2 nil)))
(terpri)
(write (cons 1 (cons 2 (cons 3 nil))))
(terpri)
(write (cons 'a (cons 'b (cons 'c nil))))
(terpri)
(write ( car (cons 'a (cons 'b (cons 'c nil)))))
(terpri)
(write ( cdr (cons 'a (cons 'b (cons 'c nil)))))

#| |#
;Lists in LISP


;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;

(write (list 1 2))
(terpri)
(write (list 'a 'b))
(terpri)
(write (list 1 nil))
(terpri)
(write (list 1 2 3))
(terpri)
(write (list 'a 'b 'c))
(terpri)
(write (list 3 4 'a (car '(b . c)) (* 4 -2)))
(terpri)
(write (list (list 'a 'b) (list 'c 'd 'e)))


;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;

(defun my-library (title author rating availability)
  (list :title title :author author :rating rating :availabilty availability)
)
(write (getf (my-library "Hunger Game" "Collins" 9 t) :title))


#| |#
;List Manipulating Functions


;

;car				
;cdr				
;cons				
;list				
;append				
;last				
;member				
;reverse			

;

;Example 3
;Create a new source code file named main.lisp and type the following code in it.
;

(write (car '(a b c d e f)))
(terpri)
(write (cdr '(a b c d e f)))
(terpri)
(write (cons 'a '(b c)))
(terpri)
(write (list 'a '(b c) '(e f)))
(terpri)
(write (append '(b c) '(e f) '(p q) '() '(g)))
(terpri)
(write (last '(a b c d (e f))))
(terpri)
(write (reverse '(a b c d (e f))))

#| |#
;Concatenation of car and cdr Functions

;Example 4
;Create a new source code file named main.lisp and type the following code in it.
;

(write (cadadr '(a (c d) (e f g))))
(terpri)
(write (caar (list (list 'a 'b) 'c)))   
(terpri)
(write (cadr (list (list 1 2) (list 3 4))))
(terpri)

#| |#
;Symbols


;Property Lists

;Example 1

;Create a new source code file named main.lisp and type the following code in it.
;

(write (setf (get 'books'title) '(Gone with the Wind)))
(terpri)
(write (setf (get 'books 'author) '(Margaret Michel)))
(terpri)
(write (setf (get 'books 'publisher) '(Warner Books)))


;Example 2

;Create a new source code file named main.lisp and type the following code in it.


(setf (get 'books 'title) '(Gone with the Wind))
(setf (get 'books 'author) '(Margaret Micheal))
(setf (get 'books 'publisher) '(Warner Books))

(write (get 'books 'title))
(terpri)
(write (get 'books 'author))
(terpri)
(write (get 'books 'publisher))


;Example 3
;Create a new source code file named main.lisp and type the following code in it.
;

(setf (get 'annie 'age) 43)
(setf (get 'annie 'job) 'accountant)
(setf (get 'annie 'sex) 'female)
(setf (get 'annie 'children) 3)

(terpri)
(write (symbol-plist 'annie))


;Example 4
;Create a new source code file named main.lisp and type the following code in it.
;

(setf (get 'annie 'age) 43)
(setf (get 'annie 'job) 'accountant)
(setf (get 'annie 'sex) 'female)
(setf (get 'annie 'children) 3)

(terpri)
(write (symbol-plist 'annie))
(remprop 'annie 'age)
(terpri)
(write (symbol-plist 'annie))


#| |#
;Vectors


;Creating Vectors

;The vector function allows you to make fixed-size vectors with specific values. It takes any number of arguments and returns a vector containing those arguments.
;

;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;

(setf v1 (vector 1 2 3 4 5))
(setf v2 #(a b c d e))
(setf v3 (vector 'p 'q 'r 's 't))

(write v1)
(terpri)
(write v2)
(terpri)
(write v3)



;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;

(setq a (make-array 5 :initial-element 0))
(setq b (make-array 5 :initial-element 2))

(dotimes (i 5)
  (setf (aref a i) i))
  
(write a)
(terpri)
(write b)
(terpri)

#| |#
;Fill Pointer

;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(setq a (make-array 5 :fill-pointer 0))
(write a)

(vector-push 'a a)
(vector-push 'b a)
(vector-push 'c a)

(terpri)
(write a)
(terpri)

(vector-push 'd a)
(vector-push 'e a)

;this will not be entered as the vector limit is 5
(vector-push 'f a)

(write a)
(terpri)

(vector-pop a)
(vector-pop a)
(vector-pop a)

(write a)


#| |#

;Set

;Implementing Sets in LISP


;Example
;Create a new source code file named main.lisp and type the following code in it.
;
; creating myset as an empty list
(defparameter *myset* ())
(adjoin 1 *myset*)
(adjoin 2 *myset*)

; adjoin did not change the original set
;so it remains same
(write *myset*)
(terpri)
(setf *myset* (adjoin 1 *myset*))
(setf *myset* (adjoin 2 *myset*))

;now the original set is changed
(write *myset*)
(terpri)

;adding an existing value
(pushnew 2 *myset*)

;no duplicate allowed
(write *myset*)
(terpri)

;pushing a new value
(pushnew 3 *myset*)
(write *myset*)
(terpri)

#| |#
;Checking Membership

;The member group of functions allows you to check whether an element is member of a set or not.

;Example
;Create a new source code file named main.lisp and type the following code in it.

(write (member 'zara '(ayan abdul zara riyan nuha)))
(terpri)
(write (member-if #'evenp '(3 7 2 5/3 'a)))
(terpri)
(write (member-if-not #'numberp '(3 7 2 5/3 'a 'b 'c)))

#| |#

;Set Union

;The union group of functions allows you to perform set union on two lists provided as arguments to these functions on the basis of a test.

;The following are the syntaxes of these functions:

;union list1 list2 &key :test :test-not :key 
;nunion list1 list2 &key :test :test-not :key
;The union function takes two lists and returns a new list containing all the elements present in either of the lists. If there are duplications, then only one copy of the member is retained in the returned list.

;The nunion function performs the same operation but may destroy the argument lists.


;Example
;Create a new source code file named main.lisp and type the following code in it.


(setq set1 (union '(a b c) '(c d e)))
(setq set2 (union '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)) :test-not #'mismatch)
)
      
(setq set3 (union '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)))
)
(write set1)
(terpri)
(write set2)
(terpri)
(write set3)

;When you execute the code, it returns the following result:

;(A B C D E)
;(#(F H) #(5 6 7) #(A B) #(G H))
;(#(A B) #(5 6 7) #(F H) #(5 6 7) #(A B) #(G H))

;Please Note
;The union function does not work as expected without :test-not #'mismatch arguments for a list of three vectors. This is because, the lists are made of cons cells and although the values look same to us apparently, the cdr part of cells does not match, so they are not exactly same to LISP interpreter/compiler. This is the reason; implementing big sets are not advised using lists. It works fine for small sets though.

;Set Intersection
;The intersection group of functions allows you to perform intersection on two lists provided as arguments to these functions on the basis of a test.

;The following are the syntaxes of these functions:

;intersection list1 list2 &key :test :test-not :key 
;nintersection list1 list2 &key :test :test-not :key
;These functions take two lists and return a new list containing all the elements present in both argument lists. If either list has duplicate entries, the redundant entries may or may not appear in the result.


;Example



(setq set1 (intersection '(a b c) '(c d e)))
(setq set2 (intersection '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)) :test-not #'mismatch)
)
      
(setq set3 (intersection '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)))
)
(write set1)
(terpri)
(write set2)
(terpri)
(write set3)

#| |#
;Set Difference

;The set-difference group of functions allows you to perform set difference on two lists provided as arguments to these functions on the basis of a test.

;The following are the syntaxes of these functions:

;set-difference list1 list2 &key :test :test-not :key 
;nset-difference list1 list2 &key :test :test-not :key
;The set-difference function returns a list of elements of the first list that do not appear in the second list.


;Example
;Create a new source code file named main.lisp and type the following code in it.


(setq set1 (set-difference '(a b c) '(c d e)))
(setq set2 (set-difference '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)) :test-not #'mismatch)
)
(setq set3 (set-difference '(#(a b) #(5 6 7) #(f h)) 
  '(#(5 6 7) #(a b) #(g h)))
)
(write set1)
(terpri)
(write set2)
(terpri)
(write set3)



#| |#
; Tree

;
;


;Tree as List of Lists

;Let us consider a tree structure made up of cons cell that form the following list of lists:
;
((1 2) (3 4) (5 6)).
;

;Diagrammatically, it could be expressed as:
;
;Tree Structure
;Tree Functions in LISP

;Although mostly you will need to write your own tree-functionalities according to your specific need, LISP provides some tree functions that you can use.
;
;Apart from all the list functions, the following functions work especially on tree structures:
;
;Function	Description
;copy-tree x &#38; optional vecp	
;tree-equal x y &#38; 
;subst new old tree &#38; key :test :test-not :key	
;nsubst new old tree &#38; key :test :test-not :key	
;sublis alist tree &#38; key :test :test-not :key	
;nsublis alist tree &#38; key :test :test-not :key	
													
;Example 1
;Create a new source code file named main.lisp and type the following code in it.

(setq lst (list '(1 2) '(3 4) '(5 6)))
(setq mylst (copy-list lst))
(setq tr (copy-tree lst))

(write lst)
(terpri)
(write mylst)
(terpri)
(write tr)


;Example 2
;Create a new source code file named main.lisp and type the following code in it.

(setq tr '((1 2 (3 4 5) ((7 8) (7 8 9)))))
(write tr)
(setq trs (subst 7 1 tr))
(terpri)
(write trs)


;Building Your Own Tree
;Let us try to build our own tree, using the list functions available in LISP.
;

;First let us create a new node that contains some data

(defun make-tree (item)
  "it creates a new node with item."
  (cons (cons item nil) nil)
)

;Next let us add a child node into the tree - it will take two tree nodes and add the second tree as the child of the first.
;

(defun add-child (tree child)
  (setf (car tree) (append (car tree) child))
  tree)

;This function will return the first child a given tree - it will take a tree node and return the first child of that node, or nil, if this node does not have any child node.
;

(defun first-child (tree)
  (if (null tree)
     nil
     (cdr (car tree))
  )
)

;This function will return the next sibling of a given node - it takes a tree node as argument, and returns a reference to the next sibling node, or nil, if the node does not have any.
;

(defun next-sibling (tree)
  (cdr tree)
)

;Lastly we need a function to return the information in a node:
;

(defun data (tree)
  (car (car tree))
)

;Example
;This example uses the above functionalities:

; Create a new source code file named main.lisp and type the following code in it.

(defun make-tree (item)
  "it creates a new node with item."
  (cons (cons item nil) nil)
)
  (defun first-child (tree)
     (if (null tree)
        nil
        (cdr (car tree))
     )
  )

(defun next-sibling (tree)
  (cdr tree)
)
(defun data (tree)
  (car (car tree))
)
(defun add-child (tree child)
  (setf (car tree) (append (car tree) child))
  tree
)

(setq tr '((1 2 (3 4 5) ((7 8) (7 8 9)))))
(setq mytree (make-tree 10))

(write (data mytree))
(terpri)
(write (first-child tr))
(terpri)
(setq newtree (add-child tr mytree))
(terpri)
(write newtree)

#| |#
;Hash table

;Sructure represents a collection of key-and-value pairs that are organized based on the hash code of the key. It uses the key to access the elements in the collection.
;
;A hash table is used when you need to access elements by using a key, and you can identify a useful key value. Each item in the hash table has a key/value pair. The key is used to access the items in the collection.
;

;Creating Hash Table in LISP

;In Common LISP, hash table is a general-purpose collection. You can use arbitrary objects as a key or indexes.;
;When you store a value in a hash table, you make a key-value pair, and store it under that key. Later you can retrieve the value from the hash table using the same key. Each key maps to a single value, although you can store a new value in a key.;
;Hash tables, in LISP, could be categorised into three types, based on the way the keys could be compared - eq, eql or equal. If the hash table is hashed on LISP objects then the keys are compared with eq or eql. If the hash table hash on tree structure, then it would be compared using equal.;
;The make-hash-table function is used for creating a hash table. Syntax for this function is:;
;make-hash-table &key :test :size :rehash-size :rehash-threshold;Where:
;The key argument provides the key.;
;The :test argument determines how keys are compared - it should have one of three values #'eq, #'eql, or #'equal, or one of the three symbols eq, eql, or equal. If not specified, eql is assumed.;
;The :size argument sets the initial size of the hash table. This should be an integer greater than zero.;
;The :rehash-size argument specifies how much to increase the size of the hash table when it becomes full. This can be an integer greater than zero, which is the number of entries to add, or it can be a floating-point number greater than 1, which is the ratio of the new size to the old size. The default value for this argument is implementation-dependent.;
;The :rehash-threshold argument specifies how full the hash table can get before it must grow. This can be an integer greater than zero and less than the :rehash-size (in which case it will be scaled whenever the table is grown), or it can be a floating-point number between zero and 1. The default value for this argument is implementation-dependent.;
;You can also call the make-hash-table function with no arguments.;
;Retrieving Items from and Adding Items into the Hash Table
;The gethash function retrieves an item from the hash table by searching for its key. If it does not find the key, then it returns nil.;
;It has the following syntax:;
;gethash key hash-table &optional default
;where:
;;key: is the associated key
;;hash-table: is the hash-table to be searched
;;default: is the value to be returned, if the entry is not found, which is nil, if not specified.
;;The gethash function actually returns two values, the second being a predicate value that is true if an entry was found, and false if no entry was found.
;;For adding an item to the hash table, you can use the setf function along with the gethash function.
;


;Example
;Create a new source code file named main.lisp and type the following code in it.

(setq empList (make-hash-table)) 
(setf (gethash '001 empList) '(Charlie Brown))
(setf (gethash '002 empList) '(Freddie Seal)) 
(write (gethash '001 empList)) 
(terpri)
(write (gethash '002 empList))  


;Example
;Create a new source code file named main.lisp and type the following code in it.

(setq empList (make-hash-table)) 
(setf (gethash '001 empList) '(Charlie Brown))
(setf (gethash '002 empList) '(Freddie Seal)) 
(setf (gethash '003 empList) '(Mark Mongoose)) 

(write (gethash '001 empList)) 
(terpri)
(write (gethash '002 empList)) 
(terpri)
(write (gethash '003 empList))  
(remhash '003 empList)
(terpri)
(write (gethash '003 empList))  

;Example
;Create a new source code file named main.lisp and type the following code in it.

(setq empList (make-hash-table)) 
(setf (gethash '001 empList) '(Charlie Brown))
(setf (gethash '002 empList) '(Freddie Seal)) 
(setf (gethash '003 empList) '(Mark Mongoose)) 

(maphash #'(lambda (k v) (format t "~a => ~a~%" k v)) empList)

#| |#
;Input & Output


;
;Input Functions

;The following table provides the most commonly used input functions of LISP:
;
;SL No.	Functions and Descriptions;1	
;read &#38; optional input-stream eof-error-p eof-value recursive-p;

;read-line &#38; optional input-stream eof-error-p eof-value recursive-p;

;listen &#38; optional input-stream;

;clear-input &#38; optional input-stream;


;parse-integer string &#38; key :start :end :radix :junk-allowed;

;Reading Input from Keyboard
;The read function is used for taking input from the keyboard. It may not take any argument.
;

;For example, consider the code snippet:
;
(write ( + 15.0 (read)))

;Example
;Create a new source code file named main.lisp and type the following code in it:

; the function AreaOfCircle
; calculates area of a circle
; when the radius is input from keyboard

(defun AreaOfCircle()
(terpri)
(princ "Enter Radius: ")
(setq radius (read))
(setq area (* 3.1416 radius radius))
(princ "Area: ")
(write area))
(AreaOfCircle)


;Example
;Create a new source code file named main.lisp and type the following code in it.

(with-input-from-string (stream "Welcome to Tutorials Point!")
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (read-char stream))
  (print (peek-char nil stream nil 'the-end))
  (values)
)

#| |#
;The Output Functions

;SL No.	Functions and Descriptions
;1	

; write object &#38; key :stream :escape :radix :base :circle :pretty :level :length :case :gensym :array
;
; write object &#38; key :stream :escape :radix :base :circle :pretty :level :length :case :gensym :array :readably :right-margin :miser-width :lines :pprint-dispatch
;

;Both write the object to the output stream specified by :stream, which defaults to the value of *standard-output*. Other values default to the corresponding global variables set for printing.
;

;2	

;prin1 object &#38; optional output-stream
;


;prin1 returns the object as its value.
;
;print prints the object with a preceding newline and followed by a space. It returns object.
;
;pprint is just like print except that the trailing space is omitted.
;
;princ is just like prin1 except that the output has no escape character
;



write-to-string object &#38; key :escape :radix :base :circle :pretty :level :length :case :gensym :array

write-to-string object &#38; key :escape :radix :base :circle :pretty :level :length :case :gensym :array :readably :right-margin :miser-width :lines :pprint-dispatch
;
;prin1-to-string object
;
;princ-to-string object
;
;The object is effectively printed and the output characters are made into a string, which is returned.
;
;4	
;write-char character &#38; optional output-stream
;
;It outputs the character to output-stream, and returns character.
;
;5	
;write-string string &#38; optional output-stream &#38; key :start :end
;
;It writes the characters of the specified substring of string to the output-stream.
;
;6	

;write-line string &#38; optional output-stream &#38; key :start :end

;Example
;Create a new source code file named main.lisp and type the following code in it.
;
;; this program inputs a numbers and doubles it
(defun DoubleNumber()
  (terpri)
  (princ "Enter Number : ")
  (setq n1 (read))
  (setq doubled (* 2.0 n1))
  (princ "The Number: ")
  (write n1)
  (terpri)
  (princ "The Number Doubled: ")
  (write doubled)
)
(DoubleNumber)

#| |#
;destination is standard output

;control-string holds the characters to be output and the printing directive.

;A format directive consists of a tilde (~), optional prefix parameters separated by commas, optional colon (:) and at-sign (@) modifiers, and a single character indicating what kind of directive this is.
;
;The prefix parameters are generally integers, notated as optionally signed decimal numbers.
;
;The following table provides brief description of the commonly used directives:
;
;Directive	Description
;~A	Is followed by ASCII arguments
;~S	Is followed by S-expressions
;~D	For decimal arguments
;~B	For binary arguments
;~O	For octal arguments
;~X	For hexadecimal arguments
;~C	For character arguments
;~F	For Fixed-format floating-point arguments.
;~E	Exponential floating-point arguments
;~$	Dollar and floating point arguments.
;~%	A new line is printed
;~*	Next argument is ignored
;~?	Indirection. The next argument must be a string, and the one after it a list.

;Example
;Let us rewrite the program calculating a circle's area:
;
;Create a new source code file named main.lisp and type the following code in it.
;
(defun AreaOfCircle()
  (terpri)
  (princ "Enter Radius: ")
  (setq radius (read))
  (setq area (* 3.1416 radius radius))
  (format t "Radius: = ~F~% Area = ~F" radius area)
)
(AreaOfCircle)


;(#|File I/O ; ..generate send datas....how to open excel from lisp


#| |#
;Opening Files


;The :direction keyword specifies whether the stream should handle input, output, or both, it takes the following values:
;           
;:input 	
;:output 	
;:io 		
;:probe 	
;The :element-type specifies the type of the unit of transaction for the stream.
;

;
;:error 			
;:new-version 		
;:rename 			
;:rename-and-delete 
;:append 			
;:supersede 		
;nil		 		

;
;:error   - it signals an error.
;:create - it creates an empty file with the specified name and then uses it.
;nil 	 -	it does not create a file or even a stream, but instead simply returns nil to indicate failure.
;The :external-format argument specifies an implementation-recognized scheme for representing characters in files.
;

;For example, you can open a file named myfile.txt stored in the /tmp folder as:
;
(open "/tmp/myfile.txt")

;Writing to and Reading from Files


;
;It has the following syntax:
;

with-open-file (stream filename {options}*)
  {declaration}* {form}*
;filename is the name of the file to be opened; it may be a string, a pathname, or a stream.

;
;The options are same as the keyword arguments to the function open.
;

;Example 1
;Create a new source code file named main.lisp and type the following code in it.
;
(with-open-file (stream "/tmp/myfile.txt" :direction :output)
  (format stream "Welcome to Tutorials Point!")
  (terpri stream)
  (format stream "This is a tutorials database")
  (terpri stream)
  (format stream "Submit your Tutorials, White Papers and Articles into our Tutorials   Directory.")
)


;Example 2
;Create a new source code file named main.lisp and type the following code in it.
;
(let ((in (open "/tmp/myfile.txt" :if-does-not-exist nil)))
  (when in
     (loop for line = (read-line in nil)
     
     while line do (format t "~a~%" line))
     (close in)
  )
)

#| |#
;Structure

;fined data type, which allows you to combine data items of different kinds.
;
;Structures are used to represent a record. Suppose you want to keep track of your books in a library. You might want to track the following attributes about each book:
;
;Title
;Author
;Subject
;Book ID

;Defining a Structure
;The defstruct macro in LISP allows you to define an abstract record structure. 
; The defstruct statement defines a new data type, with more than one member for your program.

(defstruct book 
  title 
  author 
  subject 
  book-id 
)

(setf (book-book-id book3) 100)

;Example
;Create a new source code file named main.lisp and type the following code in it.
;
(defstruct book 
  title 
  author 
  subject 
  book-id 
)
( setq book1 (make-book :title "C Programming"
  :author "Nuha Ali" 
  :subject "C-Programming Tutorial"
  :book-id "478")
)
( setq book2 (make-book :title "Telecom Billing"
  :author "Zara Ali" 
  :subject "C-Programming Tutorial"
  :book-id "501")
) 
(write book1)
(terpri)
(write book2)
(setq book3( copy-book book1))
(setf (book-book-id book3) 100) 
(terpri)
(write book3)

#| |#
;Packages

;common-lisp - it contains symbols for all the functions and variables defined.
;
;common-lisp-user - it uses the common-lisp package and all other packages with editing and debugging tools; it is called cl-user in short

;

;Package Functions in LISP


;3	

;in-package name

;find-package name
;


; renames a package.
;
;	

;list-all-packages
;
;This function returns a list of all packages that currently exist in the Lisp system.
;
;7	

;delete-package package

;it deletes a package


;Creating a LISP Package

;The defpackage function is used for creating an user defined package. It has the following syntax:
;

(defpackage :package-name
  (:use :common-lisp ...)
  (:export :symbol1 :symbol2 ...)
)

#| |#
;Using a Package



;Example
;Create a new source code file named main.lisp and type the following code in it.

(make-package :tom)
(make-package :dick)
(make-package :harry)
(in-package tom)
(defun hello () 
  (write-line "Hello! This is Tom's Tutorials Point")
)
(hello)
(in-package dick)
(defun hello () 
  (write-line "Hello! This is Dick's Tutorials Point")
)
(hello)
(in-package harry)
(defun hello () 
  (write-line "Hello! This is Harry's Tutorials Point")
)
(hello)
(in-package tom)
(hello)
(in-package dick)
(hello)
(in-package harry)
(hello)


;Example
;Create a new source code file named main.lisp and type the following code in it.
;
(make-package :tom)
(make-package :dick)
(make-package :harry)
(in-package tom)
(defun hello () 
  (write-line "Hello! This is Tom's Tutorials Point")
)
(in-package dick)
(defun hello () 
  (write-line "Hello! This is Dick's Tutorials Point")
)
(in-package harry)
(defun hello () 
  (write-line "Hello! This is Harry's Tutorials Point")
)
(in-package tom)
(hello)
(in-package dick)
(hello)
(in-package harry)
(hello)
(delete-package tom)
(in-package tom)
(hello)

#| |#
;(#|Error Handlnig


(define-condition condition-name (error)
  ((text :initarg :text :reader text))
)

;New condition objects are created with MAKE-CONDITION macro, which initializes the slots of the new condition based on the :initargs argument.
;
;In our example, the following code defines the condition:
;

(define-condition on-division-by-zero (error)
  ((message :initarg :message :reader message))
)

;Writing the Handlers - a condition handler is a code that are used for handling the condition signalled thereon. It is generally written in one of the higher level functions that call the erroring function. When a condition is signalled, the signalling mechanism searches for an appropriate handler based on the condition's class.
;

;Each handler consists of:

;

;The macro handler-case establishes a condition handler. The basic form of a handler-case:
;
(handler-case expression error-clause*)

;Where, each error clause is of the form:
;
condition-type ([var]) code)

;Restarting Phase
;

(handler-bind (binding*) form*)


;Example
;In this example, we demonstrate the above concepts by writing a function named division-function, which will create an error condition if the divisor argument is zero. We have three anonymous functions that provide three ways to come out of it - by returning a value 1, by sending a divisor 2 and recalculating, or by returning 1.
;
;Create a new source code file named main.lisp and type the following code in it.

(define-condition on-division-by-zero (error)
  ((message :initarg :message :reader message))
)
  
(defun handle-infinity ()
  (restart-case
     (let ((result 0))
        (setf result (division-function 10 0))
        (format t "Value: ~a~%" result)
     )
     (just-continue () nil)
  )
)
    
(defun division-function (value1 value2)
  (restart-case
     (if (/= value2 0)
        (/ value1 value2)
        (error 'on-division-by-zero :message "denominator is zero")
     )

     (return-zero () 0)
     (return-value (r) r)
     (recalc-using (d) (division-function value1 d))
  )
)

(defun high-level-code ()
  (handler-bind
     (
        (on-division-by-zero
           #'(lambda (c)
              (format t "error signaled: ~a~%" (message c))
              (invoke-restart 'return-zero)
           )
        )
        (handle-infinity)
     )
  )
)

(handler-bind
  (
     (on-division-by-zero
        #'(lambda (c)
           (format t "error signaled: ~a~%" (message c))
           (invoke-restart 'return-value 1)
        )
     )
  )
  (handle-infinity)
)

(handler-bind
  (
     (on-division-by-zero
        #'(lambda (c)
           (format t "error signaled: ~a~%" (message c))
           (invoke-restart 'recalc-using 2)
        )
     )
  )
  (handle-infinity)
)

(handler-bind
  (
     (on-division-by-zero
        #'(lambda (c)
           (format t "error signaled: ~a~%" (message c))
           (invoke-restart 'just-continue)
        )
     )
  )
  (handle-infinity)
)

(format t "Done."))

#| |#
;Error Signalling Functions in LISP


;Example
;In this example, the factorial function calculates factorial of a number; however, if the argument is negative, it raises an error condition.
;
;Create a new source code file named main.lisp and type the following code in it.
;

(defun factorial (x)
  (cond ((or (not (typep x 'integer)) (minusp x))
     (error "~S is a negative number." x))
     ((zerop x) 1)
     (t (* x (factorial (- x 1))))
  )
)
        
(write(factorial 5))
(terpri)
(write(factorial -1))

;When you execute the code, it returns the following resul
;|#)
;
;
;

;(#|CLOS

;vance of object-oriented programming by couple of decades. However, it object-orientation was incorporated into it at a later stage.
;

;Defining Classes
;The defclass macro allows creating user-defined classes. It establishes a class as a data type. It has the following syntax:
;

(defclass class-name (superclass-name*)
  (slot-description*)
  class-option*))

;The slots are variables that store data, or fields.


;For example, let us define a Box class, with three slots length, breadth, and height.
;

(defclass Box () 
  (length 
  breadth 
  height)
)

;Providing Access and Read/Write Control to a Slot
;Unless the slots have values that can be accessed, read or written to, classes are pretty useless.
;
;You can specify accessors for each slot when you define a class. For example, take our Box class:
;

(defclass Box ()
  ((length :accessor length)
     (breadth :accessor breadth)
     (height :accessor height)
  )
)

You can also specify separate accessor names for reading and writing a slot.


(defclass Box ()
  ((length :reader get-length :writer set-length)
     (breadth :reader get-breadth :writer set-breadth)
     (height :reader get-height :writer set-height)
  )
)
#| |#
;Creating Instance of a Class


(make-instance class {initarg value}*)

;Example
;Let us create a Box class, with three slots, length, breadth and height. We will use three slot accessors to set the values in these fields.
;
;Create a new source code file named main.lisp and type the following code in it.
;

(defclass box ()
  ((length :accessor box-length)
     (breadth :accessor box-breadth)
     (height :accessor box-height)
  )
)
(setf item (make-instance 'box))
(setf (box-length item) 10)
(setf (box-breadth item) 10)
(setf (box-height item) 5)
(format t "Length of the Box is ~d~%" (box-length item))
(format t "Breadth of the Box is ~d~%" (box-breadth item))
(format t "Height of the Box is ~d~%" (box-height item))


;Create a new source code file named main.lisp and type the following code in it.
;
(defclass box ()
  ((length :accessor box-length)
     (breadth :accessor box-breadth)
     (height :accessor box-height)
     (volume :reader volume)
  )
)

; method calculating volume   

(defmethod volume ((object box))
  (* (box-length object) (box-breadth object)(box-height object))
)

;setting the values 

(setf item (make-instance 'box))
(setf (box-length item) 10)
(setf (box-breadth item) 10)
(setf (box-height item) 5)


; displaying values


(format t "Length of the Box is ~d~%" (box-length item))
(format t "Breadth of the Box is ~d~%" (box-breadth item))
(format t "Height of the Box is ~d~%" (box-height item))
(format t "Volume of the Box is ~d~%" (volume item))

#| |#
;Inheritance

;LISP allows you to define an object in terms of another object. This is called inheritance. You can create a derived class by adding features that are new or different. The derived class inherits the functionalities of the parent class.
;
;The following example explains this:
;

;Example
;Create a new source code file named main.lisp and type the following code in it.
;

(defclass box ()
  ((length :accessor box-length)
     (breadth :accessor box-breadth)
     (height :accessor box-height)
     (volume :reader volume)
  )
)
; method calculating volume   
(defmethod volume ((object box))
  (* (box-length object) (box-breadth object)(box-height object))
)
 
;wooden-box class inherits the box class  
(defclass wooden-box (box)
((price :accessor box-price)))

;setting the values 
(setf item (make-instance 'wooden-box))
(setf (box-length item) 10)
(setf (box-breadth item) 10)
(setf (box-height item) 5)
(setf (box-price item) 1000)

; displaying values

(format t "Length of the Wooden Box is ~d~%" (box-length item))
(format t "Breadth of the Wooden Box is ~d~%" (box-breadth item))
(format t "Height of the Wooden Box is ~d~%" (box-height item))
(format t "Volume of the Wooden Box is ~d~%" (volume item))
(format t "Price of the Wooden Box is ~d~%" (box-price item))

