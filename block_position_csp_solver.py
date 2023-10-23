import sys

from constraint import *
from excel_utils import index_to_coordinate_string,coordinate_string_to_index
from config_reader import ConfigReader

class CspSolverConstraints():
    def __init__(self):
        self.constraint_func_list = []
        self.variable_pair_list = []
        self.constraint_priority_list = []
        self.max_priority_value = 0

    def __len__(self):
        return len(self.constraint_func_list)

    def get_all_constraints(self):
        return self.get_constraints_with_priority_less_or_equal_to()

    def get_constraints_with_priority_less_or_equal_to(self,priority_th=sys.maxsize):
        constraint_func_and_variable_pair_list = []
        for i in range(len(self)):
            cur_constraint_priority = self.constraint_priority_list[i]
            if cur_constraint_priority<=priority_th:
                constraint_func_and_variable_pair_list.append((self.constraint_func_list[i],self.variable_pair_list[i]))

        return constraint_func_and_variable_pair_list

    def add_constraint(self,constraint_func,variable_pair,priority=0):
        self.constraint_func_list.append(constraint_func)
        self.variable_pair_list.append(variable_pair)
        self.constraint_priority_list.append(priority)
        self.max_priority_value = max(self.max_priority_value,priority)


class BlockPositionCSPSolver():
    def __init__(self,block_size_dict):
        '''
        block_size_dict is like {'schedule_block': {'width': 10, 'height': 40}, 'rule_block': {'width': 4, 'height': 7}}
        '''
        self.problem = Problem()
        self.block_size_dict = block_size_dict
        self.width_sum, self.height_sum = self._calculate_width_and_height_sum(self.block_size_dict)
        self.constraints_manager = CspSolverConstraints()
        self._add_size_constraints()

    def _add_constraint(self,constraint_func,variable_pair,priority=0):
        self.constraints_manager.add_constraint(constraint_func,variable_pair,priority)

    def is_variable_exists(self,variable_name):
        return variable_name in self.variables()

    def constraints(self):
        return self.problem._constraints

    def variables(self):
        return self.problem._variables.keys()

    @staticmethod
    def _calculate_width_and_height_sum(block_size_dict):
        height_sum = 0
        width_sum = 0
        for block_name, size_dict in block_size_dict.items():
            if 'height' in size_dict and size_dict['height']!=-1:
                height_sum += size_dict['height']
            if 'width' in size_dict and size_dict['width']!=-1:
                width_sum += size_dict['width']
        return width_sum, height_sum

    @staticmethod
    def _variable_string_to_object_and_state(variable_string):
        splited_variable_string = variable_string.split('-')
        object = splited_variable_string[0]
        state = splited_variable_string[1]
        return object, state

    def _object_and_state_to_variable_string(self,object,state):
        return f"{object}-{state}"

    def _get_constraint_func(self,relationship,value=0):
        if relationship == 'aligned':
            return lambda a, b: a==b
        elif relationship == 'adjacent':
            return lambda a, b: a+1==b
        elif relationship == 'length':
            return lambda a, b: b-a+1 == value
        elif relationship == 'length_no_less_than':
            return lambda a, b: b-a+1 >=value
        else:
            assert False, f"Unrecognized relationship: {relationship}"
            return None

    def _get_range(self,state):
        '''
        state is a string, like start_column, start_row, end_column, or end_row.
        '''
        # TODO: return range according to column or row.
        if state == 'start_column' or state == 'end_column':
            return range(1,self.width_sum+1)
        elif state == 'start_row' or state == 'end_row':
            return range(1,self.height_sum+1)
        assert Flase, f"Unsupport state: {state}"

    def _update_domain(self,variable_name,variable_range):
        self.problem._variables[variable_name] = Domain(variable_range)

    def _add_variable(self,object,state,variable_range):
        variable_name = self._object_and_state_to_variable_string(object,state)
        self.problem.addVariable(variable_name, variable_range)
        return variable_name

    def set_variable(self,object,state,value):
        variable_name = self._object_and_state_to_variable_string(object,state)
        self._update_domain(variable_name,[value])

    def _add_width_constraint(self,object,width):
        primary_variable = self._add_variable(object,'start_column',self._get_range('start_column'))
        secondary_variable = self._add_variable(object,'end_column',self._get_range('end_column'))
        
        constraint_func = self._get_constraint_func('length',width)
        # width constraint is always the highest priority.
        self._add_constraint(constraint_func,(primary_variable,secondary_variable))

    def _add_height_constraint(self,object,height):
        primary_variable = self._add_variable(object,'start_row',self._get_range('start_row'))
        secondary_variable = self._add_variable(object,'end_row',self._get_range('end_row'))

        constraint_func = self._get_constraint_func('length',height)
        # height constraint is always the highest priority.
        self._add_constraint(constraint_func,(primary_variable,secondary_variable))

    def _add_min_height_constraint(self,object, min_height):
        primary_variable = self._add_variable(object,'start_row',self._get_range('start_row'))
        secondary_variable = self._add_variable(object,'end_row',self._get_range('end_row'))

        constraint_func = self._get_constraint_func('length_no_less_than',min_height)
        # height constraint is always the highest priority.
        self._add_constraint(constraint_func,(primary_variable,secondary_variable))

    def _add_size_constraints(self):
        for block_name, block_size in self.block_size_dict.items():
            if 'width' in block_size and block_size['width']!=-1:
                self._add_width_constraint(block_name, block_size['width'])
            if 'height' in block_size and block_size['height']!=-1:
                self._add_height_constraint(block_name, block_size['height'])
            if 'min_height' in block_size and block_size['min_height']!=-1:
                self._add_min_height_constraint(block_name, block_size['min_height'])

    def add_relative_position_constraint(self,primary_object,primary_state,secondary_object,secondary_state,relationship,priority):
        '''
        input arguments can be like:
        solver.add_relative_position_constraint('schedule_block','start_column','rule_block','start_column','aligned')
        '''
        primary_variable = self._object_and_state_to_variable_string(primary_object,primary_state)
        if not self.is_variable_exists(primary_variable):
            self._add_variable(primary_object,primary_state,self._get_range(primary_state))
        secondary_variable = self._object_and_state_to_variable_string(secondary_object,secondary_state)
        if not self.is_variable_exists(secondary_variable):
            self._add_variable(secondary_object,secondary_state,self._get_range(secondary_state))

        constraint_func = self._get_constraint_func(relationship)
        self._add_constraint(constraint_func,(primary_variable,secondary_variable),priority)

    def _solve_with_priority_less_or_equal_to(self,priority):
        # Redefine all the constraints.
        del self.problem._constraints[:]
        for constraint_func, variable_pair in self.constraints_manager.get_constraints_with_priority_less_or_equal_to(priority):
            self.problem.addConstraint(constraint_func,variable_pair)

        solution =  self.problem.getSolutions()
        return solution

    def _find_best_solution(self,solutions):
        '''
        the best solution takes the least spaces for all the blocks, so we determine the best solution by the sum of column num and row num.
        '''
        def _get_boundary_of_solution(solution):
            end_column = 0
            end_row = 0
            for variable, value in solution.items():
                if 'end_column' in variable:
                    end_column = max(end_column,value)
                elif 'end_row' in variable:
                    end_row = max(end_row,value)
            return end_column, end_row

        best_solution = None
        min_sum_of_column_num_and_row_num = sys.maxsize
        for solution in solutions:
            end_col, end_row = _get_boundary_of_solution(solution)
            if end_col+end_row < min_sum_of_column_num_and_row_num:
                best_solution = solution
        return best_solution

    def solve(self):
        solutions = None
        for priority in range(self.constraints_manager.max_priority_value, -1, -1):
            solutions = self._solve_with_priority_less_or_equal_to(priority)
            if len(solutions) >= 1:
                break
        assert solutions is not None, "Problem can not be solved."
        if len(solutions)>1:
            best_solution = self._find_best_solution(solutions)
        else:
            best_solution = solutions[0]
        return BlockPositionCSPSolver.solution_to_excel_start_coords(best_solution)

    @staticmethod
    def solution_to_excel_start_coords(solution):
        start_column_row_dict = {}
        for variable_name, value in solution.items():
            object_name, state_name = BlockPositionCSPSolver._variable_string_to_object_and_state(variable_name)
            if object_name not in start_column_row_dict:
                start_column_row_dict[object_name]={state_name:value}
            else:
                start_column_row_dict[object_name][state_name]=value
        
        # start_column_row_dict to start_coord_dict
        strat_coord_dict = {}
        for object, start_column_and_start_row in start_column_row_dict.items():
            strat_column = start_column_and_start_row['start_column']
            strat_row = start_column_and_start_row['start_row']
            strat_coord_dict[object] = index_to_coordinate_string(strat_column,strat_row)
        return strat_coord_dict

    def add_compact_constraint(self,constraint,object_list,priority=0):
        '''
        compact constraint includes: left_aligned_bottom_adjacent, right_aligned_bottom_adjacent, top_aligned_left_adjacent, bottom_aligned_left_adjacent.
        object_list is like: ['schedule_block', 'rule_block']
        The first element in the object_list will be used as object to be aligned.
        The smaller the priority value, the higher priority.
        '''
        def deal_with_left_aligned_bottom_adjacent(object_list):
            primary_object = object_list[0]
            for i in range(1,len(object_list)):
                secondary_object = object_list[i]
                self.add_relative_position_constraint(primary_object,'start_column',secondary_object,'start_column','aligned',priority)
                upper_object = object_list[i-1]
                lower_object = object_list[i]
                self.add_relative_position_constraint(upper_object,'end_row',lower_object,'start_row','adjacent',priority)
        
        def deal_with_right_aligned_bottom_adjacent(object_list):
            primary_object = object_list[0]
            for i in range(1,len(object_list)):
                secondary_object = object_list[i]
                self.add_relative_position_constraint(primary_object,'end_column',secondary_object,'end_column','aligned',priority)
                upper_object = object_list[i-1]
                lower_object = object_list[i]
                self.add_relative_position_constraint(upper_object,'end_row',lower_object,'start_row','adjacent',priority)

        def deal_with_top_aligned_left_adjacent(object_list):
            primary_object = object_list[0]
            for i in range(1,len(object_list)):
                secondary_object = object_list[i]
                self.add_relative_position_constraint(primary_object,'start_row',secondary_object,'start_row','aligned',priority)
                left_object = object_list[i-1]
                right_object = object_list[i]
                self.add_relative_position_constraint(left_object,'end_column',right_object,'start_column','adjacent',priority)

        def deal_with_bottom_aligned_left_adjacent(object_list):
            primary_object = object_list[0]
            for i in range(1,len(object_list)):
                secondary_object = object_list[i]
                self.add_relative_position_constraint(primary_object,'end_row',secondary_object,'end_row','aligned',priority)
                left_object = object_list[i-1]
                right_object = object_list[i]
                self.add_relative_position_constraint(left_object,'end_column',right_object,'start_column','adjacent',priority)

        assert len(object_list)>=2, f"object list has too less elements(length is {len(object_list)}), the object list is: {object_list}"
        if constraint == 'left_aligned_bottom_adjacent':
            deal_with_left_aligned_bottom_adjacent(object_list)
        elif constraint == 'right_aligned_bottom_adjacent':
            deal_with_right_aligned_bottom_adjacent(object_list)
        elif constraint == 'top_aligned_left_adjacent':
            deal_with_top_aligned_left_adjacent(object_list)
        elif constraint == 'bottom_aligned_left_adjacent':
            deal_with_bottom_aligned_left_adjacent(object_list)
        else:
            assert False, f"Unknown constraint type: {constraint}"


class BlockPositionCSPSolverAdaptor():
    def __init__(self,config_path):
        cr = ConfigReader(config_path)
        self.block_config = cr.get_config()
        self.block_size_dict = {}
    
    def set_block_size_base_on_template_block(self,template_block_size_dict):
        '''
        template_block_size_dict can be like: 
        template_block_size_dict = {'title_block': {'width': 14, 'height': 7}, 'theme_block': {'width': 10, 'height': 3}, 'parent_block': {'width': 10, 'height': 1}, 'child_block': {'width': 10, 'height': 1}, 'notice_block': {'width': 10, 'height': 1}, 'rule_block': {'width': 4, 'height': 7}, 'information_block': {'width': 4, 'height': 22}, 'contact_block': {'width': 4, 'height': 9}, 'project_block': {'width': 4, 'height': 3}}
        This method will do the following things:
        1. replace parent and child block with schedule block.
        2. 
        '''
        for template_block_name, template_block_size in template_block_size_dict.items():
            # schedule_block is an exception, it only exists in block to be written, and stacks vertially with parent template block and child template block. So they have the same width.
            if template_block_name == 'parent_block':
                self.block_size_dict['schedule_block'] = {'width':template_block_size['width']}
                continue
            elif template_block_name in ['notice_block','child_block']:
                continue
            self.block_size_dict[template_block_name] = template_block_size

    def set_schedule_block_height(self,schedule_height):
        self.block_size_dict['schedule_block']['height'] = schedule_height

    def _process_freedoms(self):
        for freedom in self.block_config['freedoms']:
            if freedom['relationship'] == 'min_height':
                for object in freedom['objects']:
                    # change height key into min_height.
                    self.block_size_dict[object]['min_height'] = self.block_size_dict[object].pop('height')
            else:
                assert Flase, f"Unrecognized freedom type: {freedom['relationship']}"

    def _process_coordinates(self,solver):
        for coordinate in self.block_config['coordinates']:
            object = coordinate['object']
            coord = coordinate['position']
            start_column, start_row = coordinate_string_to_index(coord)
            solver.set_variable(object,'start_column',start_column)
            solver.set_variable(object,'start_row',start_row)

    def _process_constraints(self,solver):
        for constraint in self.block_config['constraints']:
            relationship = constraint['relationship']
            objects = constraint['objects']
            priority = constraint['priority']
            solver.add_compact_constraint(relationship,objects,priority)

    def get_start_coords(self):
        self._process_freedoms()
        solver = BlockPositionCSPSolver(self.block_size_dict)
        self._process_coordinates(solver)
        self._process_constraints(solver)
        return solver.solve()

if __name__ == '__main__':
    # Test BlockPositionCSPSolver
    # block_size_dict = {'schedule_block': {'width': 10, 'height': 10}, 'rule_block': {'width': 4, 'height': 7},'project_block': {'width': 4,'min_height':3},'information_block':{'width':4, 'height':10}}
    # solver = BlockPositionCSPSolver(block_size_dict)
    # solver.set_variable('schedule_block','start_column',1)
    # solver.set_variable('schedule_block','start_row',1)
    # # solver.add_relative_position_constraint('schedule_block','start_column','rule_block','start_column','aligned')
    # # solver.add_relative_position_constraint('schedule_block','end_row','rule_block','start_row','adjacent')
    # solver.add_compact_constraint('left_aligned_bottom_adjacent',['rule_block','project_block','information_block'])
    # solver.add_compact_constraint('top_aligned_left_adjacent',['schedule_block','rule_block'])
    # solver.add_compact_constraint('bottom_aligned_left_adjacent',['schedule_block','information_block'],1)
    # print(solver.solve())

    # Test BlockPositionCSPSolverAdaptor
    solver = BlockPositionCSPSolverAdaptor('/home/lighthouse/tm_meeting_assistant/example/jabil_jouse_template_for_print/block_position_csp_config.yaml')
    template_block_size_dict = {'title_block': {'width': 14, 'height': 7}, 'theme_block': {'width': 10, 'height': 3}, 'parent_block': {'width': 10, 'height': 1}, 'child_block': {'width': 10, 'height': 1}, 'notice_block': {'width': 10, 'height': 1}, 'rule_block': {'width': 4, 'height': 7}, 'information_block': {'width': 4, 'height': 22}, 'contact_block': {'width': 4, 'height': 9}, 'project_block': {'width': 4, 'height': 3}}
    solver.set_block_size_base_on_template_block(template_block_size_dict)
    solver.set_schedule_block_height(10)
    print(solver.get_start_coords())