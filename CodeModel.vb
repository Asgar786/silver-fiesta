Imports System.Reflection
Namespace CodeModel
    Public Interface IOffset
        Property ILOffset() As Int32
    End Interface
    Public Interface IMethodWriter
        Function WriteMethodBody() As String()
        Function WriteMethodDefinition() As String
        Function WriteLocals() As String()
    End Interface
    Public Interface IFormatter
        Function ToString(ByVal FormatOption As FormatOption) As String
        Enum FormatOption
            OptionShort = 1
            OptionLong = 2
        End Enum
    End Interface
    Public Interface IBinaryExpression
        Inherits IOffset
        Property LeftOperand() As Object
        Property RightOperand() As Object
        Property [Operator]() As BinaryOperator
    End Interface
    Public Interface IUnaryExpression
        Inherits IOffset
        Property Operand() As Object
        Property [Operator]() As UnaryOperator
    End Interface
    Public Interface ICastExpression
        Inherits IOffset
        Property CastObject() As Object
        Property TargetType() As Type
    End Interface
    Public Interface IConvertExpression
        Inherits IOffset
        Property ValueToConvert() As Object
        Property TargetType() As Type
    End Interface
    Public Interface IMethodExpression
        Inherits IOffset
        ReadOnly Property MethodName() As String
        ReadOnly Property Parameters() As String()
        ReadOnly Property CalledOn() As Object
        Sub GetParameterObjects()
        Sub GetReferencedObject()
    End Interface
    Public Interface ICreateInstanceExpression
        Inherits IOffset
        ReadOnly Property Ctor() As String
        ReadOnly Property Parameters() As String()
        ReadOnly Property CalledOn() As Object
        Sub GetParameters()
        Sub GetReferencedObject()
    End Interface
    Public Interface IBranchExpression
        Inherits IOffset
        Property BranchLabel() As Int32
        Property BranchDirection() As BranchType
    End Interface
    Public Interface IConditionalBranchExpression
        Inherits IBranchExpression, IBinaryExpression
    End Interface
    'Block 
    Public Interface IBlock
        Property StartOffset() As Int32
        Property EndOffset() As Int32
        Property BlockBody() As List(Of IOffset)
    End Interface
    'Loops
    'For
    Public Interface IForNextBlock
        Inherits IBlock
        Property Initializer() As IBinaryExpression
        Property Condition() As IBinaryExpression
        Property Incrementer() As Int32
    End Interface
    Public Interface IForEachBlock
        Inherits IBlock
        Property VariableDeclaration() As Object
        Property Collection() As Object
    End Interface
    'While
    Public Interface IWhileBlock
        Inherits IBlock
        Property Binarycondition() As IBinaryExpression
        Property UnaryCondition() As IUnaryExpression
    End Interface
    'DoWhile
    Public Interface IDoWhileBlock
        Inherits IWhileBlock
    End Interface
    'Do
    Public Interface IDoBlock
        Inherits IWhileBlock
    End Interface
    'Branching
    Public Interface IIfThenBlock
        Inherits IBlock
        Property Binarycondition() As IBinaryExpression
        Property UnaryCondition() As IUnaryExpression
    End Interface
    Public Interface IElseIfBlock
        Inherits IIfThenBlock
    End Interface
    Public Interface IElseBlock
        Inherits IOffset
    End Interface
    Public Interface ISwitchCaseBlock
        Inherits IBlock
        Property comparer() As Object
        Property cases() As Dictionary(Of Object, IBlock)
    End Interface
    Public Interface IUsingBlock
        Inherits IBlock
        Property ObjectInstance() As ICreateInstanceExpression
        Property RightOperand() As Object
    End Interface
    Public Interface IBlockWriter
        Function WriteBlock(ByVal block As IBlock) As String()
        Function WriteIfThenBlock(ByVal block As IIfThenBlock) As String()
        Function WriteSwitchCaseBlock(ByVal block As ISwitchCaseBlock) As String()
        Function WriteUsingBlock(ByVal block As IUsingBlock) As String()
        Function WriteForNextBlock(ByVal block As IForNextBlock) As String()
        Function WriteForEachBlock(ByVal block As IForEachBlock) As String()
        Function WriteWhileBlock(ByVal block As IWhileBlock) As String()
        Function WriteDoWhileBlock(ByVal block As IDoWhileBlock) As String()
        Function WriteDoBlock(ByVal block As IDoBlock) As String()
    End Interface
    'Enums
    Public Enum BranchType
        BranchForward = 1
        BranchBack = 2
    End Enum
    Public Enum BinaryOperator
        ' Fields
        Add = 0
        BitwiseAnd = 12
        BitwiseExclusiveOr = 13
        BitwiseOr = 11
        BooleanAnd = 15
        BooleanOr = 14
        Divide = 3
        GreaterThan = 18
        GreaterThanOrEqual = 19
        IdentityEquality = 7
        IdentityInequality = 8
        LessThan = 16
        LessThanOrEqual = 17
        Modulus = 4
        Multiply = 2
        ShiftLeft = 5
        ShiftRight = 6
        Subtract = 1
        ValueEquality = 9
        ValueInequality = 10
    End Enum
    Public Enum UnaryOperator
        ' Fields
        BitwiseNot = 2
        BooleanNot = 1
        Negate = 0
        PostDecrement = 6
        PostIncrement = 5
        PreDecrement = 4
        PreIncrement = 3
    End Enum
End Namespace
