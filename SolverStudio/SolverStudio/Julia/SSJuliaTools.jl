# SolverStudio Julia Support Code
#
# Notes: 
# We cannot (for speed reasons) handle indices or indexed values that are booleans ("true" vs "True")
# We have to translate "nothing" (which is union{} in Julia 4) into None for Python
# We allow NaN values to be written (as occur when a model is infeasible), and define NaN=None in the Python input

function SS_WriteSolverStudioVariable(io,var)
	write(io,string(repr(var), "\r\n"))
end

function SS_WriteSolverStudioVariable(io,var::Bool)
	if var
			write(io,"True\r\n")
	else
			write(io,"False\r\n")
	end
end

try
	using JuMP
	function SS_WriteSolverStudioVariable(io, var::JuMP.JuMPContainer)
		write(io, "{")
		if ndims(var) == 1
			for it in var
				write(io, string(repr(it[1]), " : ", repr(it[2]), ", "))
			end
		else
			for it in var
				indices = it[1:end-1]
				value = it[end]
				write(io, string("(",join(map(repr,indices),","),") : ", repr(value), ", "))
			end
		end
		write(io, "}\r\n")
	end
end

try
	using JuMP
	function SS_WriteSolverStudioVariable(io, var::JuMP.JuMPArray)
		# We assume a JuMPArray will never contain "nothing", and so do not turn this into "None"
		write(io, "{")
		for it in keys(var)
			if length(it)==1
				write(io, string(repr(it[1]), " : ", repr(var[it...]), ", ")) # Output the index as a singleton, not a tuple
			else
				write(io, string(repr(it), " : ", repr(var[it...]), ", "))
			end
		end
		write(io, "}\r\n")
	end
end

function SS_WriteSolverStudioVariable(io, var::Dict)
	write(io, "{ ")
	for (index, value) in var
		write(io, string(repr(index), " : ", (value == nothing ? "None" : repr(value)), ", "))
	end
	write(io, "}\r\n")
end

function SS_WriteSolverStudioVariable(io, var::Array)
	write(io, "[ ")
	for value in var
		write(io, string( (value == nothing ? "None" : repr(value)), ", "))
	end
	write(io, "]\r\n")
end

## The following methods are called for any DataItem that is defined by SolverStudio 
## as being a dictionary, and thus needs to be written with indices

function SS_WriteIndexedSolverStudioVariable(io, var::Array)
	write(io, "{")
	if ndims(var) == 1
		for it in eachindex(var) # it will be 1, 2, 3, ...
			write(io, string(repr(it), " : ", repr(var[it]), ", ")) 
		end
	else
		for it in eachindex(var) # it will be 1, 2, 3, ...; In Julia, var[1]==var[1,1], var[2]==var[1,2], etc
			write(io, string(repr(ind2sub(var,it)), " : ", (var[it] == nothing ? "None" : repr(var[it])), ", "))
		end	
	end
	write(io, "}\r\n")
end

function SS_WriteIndexedSolverStudioVariable(io, var)
	SS_WriteSolverStudioVariable(io, var)
end

